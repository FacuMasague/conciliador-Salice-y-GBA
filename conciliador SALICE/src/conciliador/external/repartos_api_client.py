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
    _resolve_empresa_targets,
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
    empresa_id: str
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


def _receipt_empresa_id(receipt: dict[str, Any]) -> str:
    return _norm_id(
        receipt.get("_empresa_id_api_original")
        or receipt.get("empresaID")
        or receipt.get("EmpresaID")
        or receipt.get("sucursalID")
        or receipt.get("sucursal_id")
    )


def fetch_repartos_detail(
    *,
    start_date: dt.date,
    end_date: dt.date,
    empresa_filter: str | None = None,
    lookback_days: int | None = None,
) -> tuple[list[dict[str, Any]], list[str]]:
    """Obtiene fojas de Salice/Alarcón sin abortar por fallos individuales."""
    lookback = _lookback_days() if lookback_days is None else min(max(int(lookback_days), 1), 60)
    fetch_from = start_date - dt.timedelta(days=lookback)
    base = _base_url("RECEIPTS_API")
    headers_root = _headers_base("RECEIPTS_API")
    targets = _resolve_empresa_targets(empresa_filter)

    selected: list[tuple[str, dict[str, Any], dict[str, str]]] = []
    list_failures = 0
    warnings: list[str] = []
    for target in targets:
        try:
            headers = _build_auth_headers_for_empresa(
                base=base,
                headers_root=headers_root,
                empresa_id=str(target),
                drop_sucursal=False,
            )
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
                if payload.get("success") is False:
                    raise ExternalProviderError(
                        "receipts", f"Repartos/GetList devolvió error para empresaID={target}"
                    )
                rows = payload.get("fojasReparto") or []
                for row in rows:
                    if not isinstance(row, dict):
                        continue
                    day = _parse_date(row.get("desde"))
                    if day is not None and fetch_from <= day <= end_date and _norm_id(row.get("fojaID")):
                        selected.append((str(target), row, headers))
                pagination = payload.get("paginacion") or {}
                try:
                    total_pages = int(pagination.get("totalPaginas") or page)
                except Exception:
                    total_pages = page
                if page >= total_pages:
                    break
                page += 1
        except Exception as exc:
            list_failures += 1
            warnings.append(f"Fleteros: Repartos/GetList empresaID={target} fue omitido: {exc}")

    def _get_item(item: tuple[str, dict[str, Any], dict[str, str]]) -> dict[str, Any]:
        target, row, headers = item
        empresa_id = _norm_id(row.get("empresaID")) or target
        query = urlencode({"empresaID": empresa_id, "fojaID": _norm_id(row.get("fojaID"))})
        payload, _ = _http_json(
            f"{base}/api/Ventas/Repartos/GetItem?{query}",
            method="GET",
            headers=headers,
        )
        if payload.get("success") is False:
            raise ExternalProviderError(
                "receipts",
                f"Repartos/GetItem devolvió error para fojaID={_norm_id(row.get('fojaID'))}",
            )
        detail = payload.get("fojaReparto") or {}
        if not isinstance(detail, dict):
            return {}
        detail.setdefault("empresaID", empresa_id)
        return detail

    # Si GESI aplica el límite de 60 solicitudes/minuto, priorizamos las fojas
    # más cercanas al lote actual antes que el extremo viejo del lookback.
    selected.sort(
        key=lambda item: (_parse_date(item[1].get("desde")) or dt.date.min, _norm_id(item[1].get("fojaID"))),
        reverse=True,
    )

    details: list[dict[str, Any]] = []
    failed = 0
    rate_limited = False
    with ThreadPoolExecutor(max_workers=_concurrency()) as executor:
        future_map = {executor.submit(_get_item, item): item for item in selected}
        for future in as_completed(future_map):
            try:
                detail = future.result()
                if detail:
                    details.append(detail)
            except ExternalProviderError as exc:
                failed += 1
                if _is_rate_limited_error(exc):
                    rate_limited = True
            except Exception:
                failed += 1

    warnings.insert(
        0,
        f"Fleteros: se consultaron {len(details)} fojas de reparto "
        f"({lookback} días de historial; empresaID={','.join(targets)}).",
    )
    if list_failures == len(targets) and targets:
        warnings.append("Fleteros: el endpoint de repartos falló completamente; los fleteros quedan vacíos salvo IDs reales del recibo.")
    if failed:
        warnings.append(f"Fleteros: {failed} fojas no pudieron leerse y fueron omitidas.")
    if rate_limited:
        warnings.append("Fleteros: GESI aplicó un límite de solicitudes al leer algunas fojas.")
    return details, warnings


def _build_candidate_index(
    fojas: list[dict[str, Any]],
) -> dict[tuple[str, str], list[_RouteCandidate]]:
    index: dict[tuple[str, str], list[_RouteCandidate]] = defaultdict(list)
    for foja in fojas:
        day = _parse_date(foja.get("desde"))
        empresa_id = _norm_id(foja.get("empresaID") or foja.get("EmpresaID"))
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
            index[(empresa_id, client_id)].append(
                _RouteCandidate(
                    date=day,
                    empresa_id=empresa_id,
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
    candidates: list[_RouteCandidate], amount: int
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


def _feature_key(features: dict[str, str]) -> tuple[str, str, str, str] | None:
    vendor = str(features.get("vendedor_id") or "")
    zone = str(features.get("zona_id") or "")
    subzone = str(features.get("subzona_id") or "")
    if vendor and zone and subzone:
        return "vendor_zone_subzone", vendor, zone, subzone
    return None


def match_receipts_to_fleteros(
    comprobantes: list[dict[str, Any]],
    fojas: list[dict[str, Any]],
    client_features: dict[str, dict[str, str]],
    *,
    lookback_days: int | None = None,
) -> tuple[dict[int, FleteroMatch], list[str]]:
    """Resuelve sólo coincidencias conservadoras; nunca devuelve el vendedor."""
    lookback = _lookback_days() if lookback_days is None else min(max(int(lookback_days), 1), 60)
    index = _build_candidate_index(fojas)
    strong: dict[int, FleteroMatch] = {}
    pending: dict[int, tuple[list[_RouteCandidate], dt.date, str]] = {}
    reason_counts: Counter[str] = Counter()

    for position, receipt in enumerate(comprobantes):
        client_id = _norm_id(receipt.get("clienteID"))
        receipt_date = _parse_date(receipt.get("fechaDeEmision"))
        empresa_id = _receipt_empresa_id(receipt)
        if not client_id or receipt_date is None:
            pending[position] = ([], receipt_date or dt.date.min, client_id)
            continue
        earliest = receipt_date - dt.timedelta(days=lookback)
        candidates = [
            candidate
            for candidate in index.get((empresa_id, client_id), [])
            if earliest <= candidate.date <= receipt_date
        ]
        candidate, source = _strong_candidate(candidates, _money_cents(receipt.get("importeTotal")))
        if candidate is not None:
            strong[position] = FleteroMatch(candidate.repartidor_id, candidate.nombre, source)
            reason_counts[source] += 1
        else:
            pending[position] = (candidates, receipt_date, client_id)

    # El respaldo comercial se aprende exclusivamente de coincidencias exactas
    # del mismo lote (empresa + fecha). Nunca se transforma vendedorID en salida.
    profiles: dict[
        tuple[str, dt.date, tuple[str, str, str, str]],
        dict[tuple[str, str], set[str]],
    ] = defaultdict(lambda: defaultdict(set))
    for position, match in strong.items():
        receipt = comprobantes[position]
        client_id = _norm_id(receipt.get("clienteID"))
        receipt_date = _parse_date(receipt.get("fechaDeEmision"))
        feature = _feature_key(client_features.get(client_id, {}))
        if receipt_date is None or not client_id or feature is None:
            continue
        profiles[(_receipt_empresa_id(receipt), receipt_date, feature)][
            (match.repartidor_id, match.nombre)
        ].add(client_id)

    result = dict(strong)
    for position, (candidates, receipt_date, client_id) in pending.items():
        receipt = comprobantes[position]
        empresa_id = _receipt_empresa_id(receipt)

        # Si todas las rutas recientes del cliente coinciden, la evidencia no es
        # contradictoria y también cubre pagos parciales.
        if candidates:
            unique = _unique_route(candidates)
            if unique is not None:
                match = FleteroMatch(unique.repartidor_id, unique.nombre, "unique_client_route")
                result[position] = match
                reason_counts[match.source] += 1
            continue

        # Sólo clientes ausentes de toda foja pueden aprender el perfil. Exigimos
        # dos clientes distintos con match fuerte en el mismo lote.
        feature = _feature_key(client_features.get(client_id, {}))
        if feature is None:
            continue
        routes = profiles.get((empresa_id, receipt_date, feature), {})
        eligible = [route for route, clients in routes.items() if len(clients) >= 2]
        if len(eligible) == 1:
            route = eligible[0]
            match = FleteroMatch(route[0], route[1], "batch_vendor_zone_subzone")
            result[position] = match
            reason_counts[match.source] += 1

    total = len(comprobantes)
    assigned = len(result)
    warnings = [f"Fleteros: {assigned}/{total} recibos identificados por fojas de reparto."]
    details = ", ".join(f"{key}={value}" for key, value in sorted(reason_counts.items()))
    if details:
        warnings.append(f"Fleteros: fuentes de asignación: {details}.")
    if assigned < total:
        warnings.append(
            f"Fleteros: {total - assigned} recibos quedaron sin identificar por falta de evidencia confiable."
        )
    return result, warnings


def resolve_fleteros(
    comprobantes: list[dict[str, Any]],
    client_features: dict[str, dict[str, str]],
    *,
    empresa_filter: str | None = None,
    start_date: dt.date | None = None,
    end_date: dt.date | None = None,
) -> tuple[dict[int, FleteroMatch], list[str]]:
    dates = [
        parsed
        for row in comprobantes
        if (parsed := _parse_date(row.get("fechaDeEmision"))) is not None
    ]
    if not dates:
        return {}, ["Fleteros: los recibos no contienen fechas válidas para consultar fojas."]
    try:
        fojas, warnings = fetch_repartos_detail(
            start_date=start_date or min(dates),
            end_date=end_date or max(dates),
            empresa_filter=empresa_filter,
        )
    except Exception as exc:
        return {}, [
            f"Fleteros: no se pudieron consultar las fojas de reparto; los valores quedan vacíos: {exc}"
        ]
    matches, match_warnings = match_receipts_to_fleteros(comprobantes, fojas, client_features)
    return matches, [*warnings, *match_warnings]
