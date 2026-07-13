from __future__ import annotations

import datetime as dt
import os
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any
from urllib.parse import urlencode

from .errors import ExternalProviderError
from .receipts_api_client import (
    _base_url,
    _build_auth_headers_for_empresa,
    _headers_base,
    _http_json,
    _is_rate_limited_error,
)


@dataclass(frozen=True)
class FleteroMatch:
    repartidor_id: str
    nombre: str
    source: str

    @property
    def label(self) -> str:
        if self.repartidor_id and self.nombre:
            return f"{self.repartidor_id} - {self.nombre}"
        return self.nombre or self.repartidor_id


@dataclass(frozen=True)
class _RouteCandidate:
    date: dt.date
    foja_id: str
    repartidor_id: str
    nombre: str
    invoice_amounts: tuple[int, ...]

    @property
    def route_key(self) -> tuple[str, str]:
        return self.repartidor_id, self.nombre

    @property
    def invoice_sum(self) -> int:
        return sum(self.invoice_amounts)


def _norm_id(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        return str(int(float(text.replace(",", "."))))
    except Exception:
        return "".join(ch for ch in text if ch.isdigit()).lstrip("0") or ""


def _parse_date(value: object) -> dt.date | None:
    text = str(value or "").strip()[:10]
    try:
        return dt.date.fromisoformat(text)
    except Exception:
        return None


def _money_cents(value: object) -> int:
    try:
        amount = Decimal(str(value or "0").replace(",", "."))
    except (InvalidOperation, ValueError):
        return 0
    return int((amount * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _lookback_days() -> int:
    raw = str(os.getenv("RECEIPTS_API_REPARTOS_LOOKBACK_DAYS", "21") or "21").strip()
    try:
        value = int(raw)
    except Exception:
        value = 21
    return min(max(value, 1), 60)


def _concurrency() -> int:
    raw = str(os.getenv("RECEIPTS_API_REPARTOS_CONCURRENCY", "6") or "6").strip()
    try:
        value = int(raw)
    except Exception:
        value = 6
    return min(max(value, 1), 12)


def fetch_repartos_detail(
    *,
    start_date: dt.date,
    end_date: dt.date,
    lookback_days: int | None = None,
) -> tuple[list[dict[str, Any]], list[str]]:
    """Obtiene las fojas que pueden respaldar los recibos del rango solicitado."""
    lookback = _lookback_days() if lookback_days is None else min(max(int(lookback_days), 1), 60)
    fetch_from = start_date - dt.timedelta(days=lookback)
    base = _base_url("RECEIPTS_API")
    headers = _build_auth_headers_for_empresa(
        base=base,
        headers_root=_headers_base("RECEIPTS_API"),
        empresa_id="2",
        drop_sucursal=False,
    )

    list_rows: list[dict[str, Any]] = []
    page = 1
    while True:
        query = urlencode({
            "fechaDesde": fetch_from.isoformat(),
            "pageNumber": str(page),
            "pageSize": "500",
        })
        payload, _ = _http_json(
            f"{base}/api/Ventas/Repartos/GetList?{query}",
            method="GET",
            headers=headers,
        )
        rows = payload.get("fojasReparto") or []
        list_rows.extend(row for row in rows if isinstance(row, dict))
        pagination = payload.get("paginacion") or {}
        try:
            total_pages = int(pagination.get("totalPaginas") or page)
        except Exception:
            total_pages = page
        if page >= total_pages:
            break
        page += 1

    selected = [
        row for row in list_rows
        if (day := _parse_date(row.get("desde"))) is not None
        and fetch_from <= day <= end_date
        and _norm_id(row.get("fojaID"))
    ]

    def _get_item(row: dict[str, Any]) -> dict[str, Any]:
        query = urlencode({
            "empresaID": _norm_id(row.get("empresaID")) or "2",
            "fojaID": _norm_id(row.get("fojaID")),
        })
        payload, _ = _http_json(
            f"{base}/api/Ventas/Repartos/GetItem?{query}",
            method="GET",
            headers=headers,
        )
        item = payload.get("fojaReparto") or {}
        return item if isinstance(item, dict) else {}

    details: list[dict[str, Any]] = []
    failed = 0
    rate_limited = False
    with ThreadPoolExecutor(max_workers=_concurrency()) as executor:
        future_map = {executor.submit(_get_item, row): row for row in selected}
        for future in as_completed(future_map):
            try:
                item = future.result()
                if item:
                    details.append(item)
            except ExternalProviderError as exc:
                failed += 1
                if _is_rate_limited_error(exc):
                    rate_limited = True
            except Exception:
                failed += 1

    warnings = [
        f"Fleteros GBA: se consultaron {len(details)} fojas de reparto "
        f"({lookback} días de historial)."
    ]
    if failed:
        warnings.append(f"Fleteros GBA: {failed} fojas no pudieron leerse y fueron omitidas.")
    if rate_limited:
        warnings.append("Fleteros GBA: GESI aplicó un límite de solicitudes al leer algunas fojas.")
    return details, warnings


def _build_candidate_index(fojas: list[dict[str, Any]]) -> dict[str, list[_RouteCandidate]]:
    index: dict[str, list[_RouteCandidate]] = defaultdict(list)
    for foja in fojas:
        day = _parse_date(foja.get("desde"))
        repartidor_id = _norm_id(foja.get("repartidorID"))
        nombre = str(foja.get("descripcionRepartidor") or "").strip()
        if day is None or not (repartidor_id or nombre):
            continue
        amounts_by_client: dict[str, list[int]] = defaultdict(list)
        for invoice in foja.get("listaDeComprobantesAsignados") or []:
            if not isinstance(invoice, dict):
                continue
            client_id = _norm_id(invoice.get("clienteID"))
            if client_id:
                amounts_by_client[client_id].append(_money_cents(invoice.get("importeTotal")))
        for client_id, amounts in amounts_by_client.items():
            index[client_id].append(
                _RouteCandidate(
                    date=day,
                    foja_id=_norm_id(foja.get("fojaID")),
                    repartidor_id=repartidor_id,
                    nombre=nombre,
                    invoice_amounts=tuple(amounts),
                )
            )
    return index


def _subset_matches(values: tuple[int, ...], target: int) -> bool:
    if target <= 0 or not values:
        return False
    possible = {0}
    for amount in values:
        if amount <= 0:
            continue
        possible.update(total + amount for total in tuple(possible) if total + amount <= target)
        if target in possible:
            return True
    return False


def _unique_route(candidates: list[_RouteCandidate]) -> _RouteCandidate | None:
    routes = {candidate.route_key for candidate in candidates}
    if len(routes) != 1:
        return None
    return max(candidates, key=lambda candidate: candidate.date)


def _strong_candidate(
    candidates: list[_RouteCandidate],
    amount: int,
) -> tuple[_RouteCandidate | None, str]:
    exact = [candidate for candidate in candidates if amount in candidate.invoice_amounts]
    if exact:
        latest_day = max(candidate.date for candidate in exact)
        unique = _unique_route([candidate for candidate in exact if candidate.date == latest_day])
        if unique is not None:
            return unique, "invoice_exact"

    summed = [candidate for candidate in candidates if candidate.invoice_sum == amount]
    unique = _unique_route(summed)
    if unique is not None:
        return unique, "foja_sum_exact"

    subset = [candidate for candidate in candidates if _subset_matches(candidate.invoice_amounts, amount)]
    unique = _unique_route(subset)
    if unique is not None:
        return unique, "invoice_subset_exact"

    return None, ""


def _feature_keys(features: dict[str, str]) -> list[tuple[str, ...]]:
    vendor = str(features.get("vendedor_id") or "")
    zone = str(features.get("zona_id") or "")
    subzone = str(features.get("subzona_id") or "")
    keys: list[tuple[str, ...]] = []
    if vendor and zone and subzone:
        keys.append(("vendor_zone_subzone", vendor, zone, subzone))
    if zone and subzone:
        keys.append(("zone_subzone", zone, subzone))
    return keys


def match_receipts_to_fleteros(
    comprobantes: list[dict[str, Any]],
    fojas: list[dict[str, Any]],
    client_features: dict[str, dict[str, str]],
    *,
    lookback_days: int | None = None,
) -> tuple[dict[int, FleteroMatch], list[str]]:
    """Vincula recibos con fleteros usando fojas y señales aprendidas del mismo lote."""
    lookback = _lookback_days() if lookback_days is None else min(max(int(lookback_days), 1), 60)
    index = _build_candidate_index(fojas)
    strong: dict[int, FleteroMatch] = {}
    pending: dict[int, tuple[list[_RouteCandidate], dt.date, str]] = {}
    reason_counts: Counter[str] = Counter()

    for position, receipt in enumerate(comprobantes):
        client_id = _norm_id(receipt.get("clienteID"))
        receipt_date = _parse_date(receipt.get("fechaDeEmision"))
        if not client_id or receipt_date is None:
            pending[position] = ([], receipt_date or dt.date.min, client_id)
            continue
        earliest = receipt_date - dt.timedelta(days=lookback)
        candidates = [
            candidate for candidate in index.get(client_id, [])
            if earliest <= candidate.date <= receipt_date
        ]
        candidate, source = _strong_candidate(candidates, _money_cents(receipt.get("importeTotal")))
        if candidate is not None:
            strong[position] = FleteroMatch(candidate.repartidor_id, candidate.nombre, source)
            reason_counts[source] += 1
        else:
            pending[position] = (candidates, receipt_date, client_id)

    # Aprende perfiles exclusivamente de coincidencias de foja confiables del
    # mismo lote/fecha. El vendedor nunca se muestra: sólo ayuda a distinguir
    # rutas cuando GESI dejó el repartidorID del recibo en cero.
    profiles: dict[tuple[dt.date, tuple[str, ...]], dict[tuple[str, str], set[str]]] = defaultdict(
        lambda: defaultdict(set)
    )
    for position, match in strong.items():
        receipt = comprobantes[position]
        receipt_date = _parse_date(receipt.get("fechaDeEmision"))
        client_id = _norm_id(receipt.get("clienteID"))
        if receipt_date is None or not client_id:
            continue
        for feature_key in _feature_keys(client_features.get(client_id, {})):
            profiles[(receipt_date, feature_key)][(match.repartidor_id, match.nombre)].add(client_id)

    result = dict(strong)
    for position, (candidates, receipt_date, client_id) in pending.items():
        receipt = comprobantes[position]
        import_code = str(receipt.get("codigoDeImportacion") or "").strip().upper()
        profile_inferred: FleteroMatch | None = None
        if import_code.startswith("PMCBR_"):
            for feature_key in _feature_keys(client_features.get(client_id, {})):
                routes = profiles.get((receipt_date, feature_key), {})
                eligible = [
                    route for route, clients in routes.items()
                    if len(clients) >= 2
                ]
                if len(eligible) == 1:
                    route = eligible[0]
                    profile_inferred = FleteroMatch(
                        route[0],
                        route[1],
                        f"batch_{feature_key[0]}",
                    )
                    break

        previous = strong.get(position - 1)
        following = strong.get(position + 1)
        if (
            candidates
            and
            import_code.startswith("PMCBR_")
            and previous is not None
            and following is not None
            and (previous.repartidor_id, previous.nombre) == (following.repartidor_id, following.nombre)
            and profile_inferred is not None
            and (previous.repartidor_id, previous.nombre)
            == (profile_inferred.repartidor_id, profile_inferred.nombre)
            and _parse_date(comprobantes[position - 1].get("fechaDeEmision")) == receipt_date
            and _parse_date(comprobantes[position + 1].get("fechaDeEmision")) == receipt_date
        ):
            inferred = FleteroMatch(
                previous.repartidor_id,
                previous.nombre,
                "batch_neighbor_profile",
            )
            result[position] = inferred
            reason_counts[inferred.source] += 1
            continue

        # El perfil sólo completa clientes ausentes de las fojas. Si hay historial,
        # aunque sea parcial, no debe ser reemplazado por una zona genérica.
        if not candidates and profile_inferred is not None:
            result[position] = profile_inferred
            reason_counts[profile_inferred.source] += 1
            continue

        if candidates:
            latest_day = max(candidate.date for candidate in candidates)
            latest = [candidate for candidate in candidates if candidate.date == latest_day]
            candidate = _unique_route(latest)
            if candidate is not None:
                match = FleteroMatch(
                    candidate.repartidor_id,
                    candidate.nombre,
                    "latest_client_route",
                )
                result[position] = match
                reason_counts[match.source] += 1

    total = len(comprobantes)
    assigned = len(result)
    details = ", ".join(f"{key}={value}" for key, value in sorted(reason_counts.items()))
    warnings = [f"Fleteros GBA: {assigned}/{total} recibos identificados por fojas de reparto."]
    if details:
        warnings.append(f"Fleteros GBA: fuentes de asignación: {details}.")
    if assigned < total:
        warnings.append(
            f"Fleteros GBA: {total - assigned} recibos quedaron sin fletero porque GESI no aportó una relación confiable."
        )
    return result, warnings


def resolve_gba_fleteros(
    comprobantes: list[dict[str, Any]],
    client_features: dict[str, dict[str, str]],
    *,
    start_date: dt.date | None = None,
    end_date: dt.date | None = None,
) -> tuple[dict[int, FleteroMatch], list[str]]:
    dates = [
        parsed for row in comprobantes
        if (parsed := _parse_date(row.get("fechaDeEmision"))) is not None
    ]
    if not dates:
        return {}, ["Fleteros GBA: los recibos no contienen fechas válidas para consultar fojas."]
    actual_start = start_date or min(dates)
    actual_end = end_date or max(dates)
    try:
        fojas, warnings = fetch_repartos_detail(start_date=actual_start, end_date=actual_end)
    except Exception as exc:
        return {}, [f"No se pudieron consultar las fojas de reparto para obtener fleteros: {exc}"]
    matches, match_warnings = match_receipts_to_fleteros(
        comprobantes,
        fojas,
        client_features,
    )
    return matches, [*warnings, *match_warnings]
