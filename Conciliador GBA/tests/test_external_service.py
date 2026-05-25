from __future__ import annotations

import pytest

from src.conciliador.external.errors import ExternalSchemaError
from src.conciliador.external.service import fetch_cliente_cuit_map, fetch_receipts_and_payments


class _Resp:
    def __init__(self, payload, request_id="req-1", warnings=None):
        self.payload = payload
        self.request_id = request_id
        self.warnings = warnings or []


def test_fetch_receipts_and_payments_maps_nested_receipts(monkeypatch):
    payload = {
        "receipts": [
            {
                "empresa": "SALICE",
                "nro_recibo": "68734",
                "nro_cliente": "33119",
                "cliente_nombre": "Fernandez Leandro Javier",
                "vendedor": "211 - Matias Carricart",
                "payments": [
                    {
                        "medio_pago": "TRANSFERENCIA",
                        "fecha_pago": "2026-02-18",
                        "importe_pago": "88200,00",
                    }
                ],
            }
        ]
    }

    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-rec"),
    )

    payments, meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].empresa == "SALICE"
    assert payments[0].nro_recibo == "68734"
    assert payments[0].importe_pago == 88200.0
    assert meta["api_request_id"] == "r-rec"


def test_fetch_receipts_and_payments_raises_on_invalid_schema(monkeypatch):
    payload = {"receipts": [{"empresa": "SALICE", "payments": [{}]}]}

    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload),
    )

    with pytest.raises(ExternalSchemaError):
        fetch_receipts_and_payments(60, None)


def test_fetch_cliente_cuit_map_maps_entries(monkeypatch):
    payload = {
        "entries": [
            {"nro_cliente": "33119", "cuit": "20-12345678-9"},
            {"nro_cliente": "33119", "cuit": "20123456789"},  # duplicate cliente ignored
            {"cliente": "30424", "numero_documento": "27-11111111-1"},
        ]
    }

    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_padron_payload",
        lambda empresa_filter=None, cliente_ids=None: _Resp(payload, request_id="r-pad"),
    )

    m, meta = fetch_cliente_cuit_map(None)
    assert m["33119"] == "20123456789"
    assert m["30424"] == "27111111111"
    assert meta["api_request_id"] == "r-pad"


def test_fetch_cliente_cuit_map_skips_invalid_rows_if_some_valid(monkeypatch):
    payload = {
        "clientes": [
            {"clienteID": 100, "numeroDeDocumento": "dni-sin-cuit"},
            {"clienteID": 101, "numeroDeDocumento": "20-12345678-9"},
            {"clienteID": "", "numeroDeDocumento": "20-99999999-9"},
        ]
    }

    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_padron_payload",
        lambda empresa_filter=None, cliente_ids=None: _Resp(payload, request_id="r-pad-2"),
    )

    m, meta = fetch_cliente_cuit_map(None)
    assert m["101"] == "20123456789"
    assert any("filas inválidas" in w for w in meta.get("external_warnings", []))


def test_fetch_cliente_cuit_map_forwards_target_cliente_ids(monkeypatch):
    seen = {}
    payload = {"clientes": [{"clienteID": 101, "numeroDeDocumento": "20-12345678-9"}]}

    def _fake_fetch_padron_payload(empresa_filter=None, cliente_ids=None):
        seen["empresa_filter"] = empresa_filter
        seen["cliente_ids"] = cliente_ids
        return _Resp(payload, request_id="r-pad-3")

    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_padron_payload",
        _fake_fetch_padron_payload,
    )

    m, _meta = fetch_cliente_cuit_map("GBA", cliente_ids=["101", "202"])
    assert m["101"] == "20123456789"
    assert seen == {"empresa_filter": "GBA", "cliente_ids": ["101", "202"]}


def test_fetch_receipts_and_payments_maps_gesi_comprobantes(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 2,
                "numero": 70001,
                "nro_recibo_pm": "PMCBR_69016",
                "codigoDeImportacion": "PMCBR_69016",
                "clienteID": 33119,
                "razonSocial": "Fernandez Leandro Javier",
                "vendedorID": 211,
                "formaDePagoID": 5,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 88200.0,
                "notas": "cobro",
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 5, "descripcion": "Transferencia Bancaria"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "2:SALICE,3:ALARCON")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi"),
    )

    payments, meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].empresa == "SALICE"
    assert payments[0].medio_pago == "Transferencia Bancaria"
    assert payments[0].nro_cliente == "33119"
    assert payments[0].nro_recibo == "70001"
    assert meta["api_request_id"] == "r-gesi"


def test_fetch_receipts_and_payments_gesi_uses_forma_pago_map(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 2,
                "numero": 1,
                "nro_recibo_pm": "PMCBR_70001",
                "codigoDeImportacion": "PMCBR_70001",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "vendedorID": 1,
                "formaDePagoID": 9,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 9, "descripcion": "Transferencia Bancaria"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "2:SALICE")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-2"),
    )

    payments, _meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].nro_recibo == "1"
    assert payments[0].medio_pago == "Transferencia Bancaria"


def test_fetch_receipts_and_payments_gesi_unknown_medio_is_kept_as_sin_medio(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 2,
                "numero": 1,
                "nro_recibo_pm": "PMCBR_70011",
                "codigoDeImportacion": "PMCBR_70011",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "vendedorID": 1,
                "formaDePagoID": 77,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 77, "descripcion": "OTRO"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "2:SALICE")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-3"),
    )

    payments, meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].medio_pago == "OTRO"
    assert meta["payments_count"] == 1
    assert meta["medio_bancarizable_stats"]["UNKNOWN"] == 1


def test_fetch_receipts_and_payments_gesi_reads_medio_from_comprobante_desc(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 2,
                "numero": 1,
                "nro_recibo_pm": "PMCBR_70012",
                "codigoDeImportacion": "PMCBR_70012",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "vendedorID": 1,
                "formaDePagoID": 77,
                "formaDePago": "Transferencia Bancaria",
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 77, "descripcion": "OTRO"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "2:SALICE")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-4"),
    )

    payments, _meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].medio_pago == "Transferencia Bancaria"


def test_fetch_receipts_and_payments_gesi_maps_empresa_from_empresas_master(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 7,
                "numero": 1,
                "nro_recibo_pm": "PMCBR_70013",
                "codigoDeImportacion": "PMCBR_70013",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "vendedorID": 1,
                "formaDePagoID": 5,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 5, "descripcion": "Transferencia Bancaria"},
        ],
        "empresas": [
            {"empresaID": 7, "descripcion": "ALARCON S.R.L."},
        ],
    }

    monkeypatch.delenv("RECEIPTS_API_EMPRESA_MAP", raising=False)
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-5"),
    )

    payments, _meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].empresa == "ALARCON"


def test_fetch_receipts_and_payments_gesi_prefers_vendor_name_when_available(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 2,
                "numero": 70021,
                "nro_recibo_pm": "PMCBR_70021",
                "codigoDeImportacion": "PMCBR_70021",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "vendedorID": 211,
                "nombreVendedor": "Matias Carricart",
                "formaDePagoID": 5,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 5, "descripcion": "Transferencia Bancaria"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "2:SALICE")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-vendor"),
    )

    payments, _meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].vendedor == "211 - Matias Carricart"


def test_fetch_receipts_and_payments_gesi_reads_vendor_from_datos_clientes(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 2,
                "numero": 70022,
                "nro_recibo_pm": "PMCBR_70022",
                "codigoDeImportacion": "PMCBR_70022",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "formaDePagoID": 5,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
                "datosClientes": {
                    "VendedorDelClienteID": 345,
                },
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 5, "descripcion": "Transferencia Bancaria"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "2:SALICE")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-datos-clientes"),
    )

    payments, _meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].vendedor == "345"


def test_fetch_receipts_and_payments_gesi_reads_nested_vendor_name_variants(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 2,
                "numero": 70023,
                "nro_recibo_pm": "PMCBR_70023",
                "codigoDeImportacion": "PMCBR_70023",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "formaDePagoID": 5,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
                "extra": {
                    "datosVendedor": {
                        "vendedorID": 211,
                        "descripcionVendedor": "Matias Carricart",
                    }
                },
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 5, "descripcion": "Transferencia Bancaria"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "2:SALICE")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-nested-vendor"),
    )

    payments, _meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].vendedor == "211 - Matias Carricart"


def test_fetch_receipts_and_payments_gesi_reads_medio_from_canal(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 2,
                "numero": 1,
                "nro_recibo_pm": "PMCBR_70014",
                "codigoDeImportacion": "PMCBR_70014",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "vendedorID": 1,
                "formaDePagoID": 9,
                "canalDeCobro": "Mercado Pago",
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 9, "descripcion": "Transferencia Bancaria"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "2:SALICE")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-6"),
    )

    payments, _meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].medio_pago == "Mercado Pago"


def test_fetch_receipts_and_payments_gesi_prefers_empresa_target_hint_over_env_map(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 2,
                "_empresa_target_name": "ALARCON",
                "numero": 1,
                "nro_recibo_pm": "PMCBR_70015",
                "codigoDeImportacion": "PMCBR_70015",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "vendedorID": 1,
                "formaDePagoID": 5,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 5, "descripcion": "Transferencia Bancaria"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "2:SALICE")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-7"),
    )

    payments, _meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].empresa == "ALARCON"


def test_fetch_receipts_and_payments_gesi_empresa_id_6_maps_to_alarcon(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 6,
                "numero": 1,
                "nro_recibo_pm": "PMCBR_70016",
                "codigoDeImportacion": "PMCBR_70016",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "vendedorID": 1,
                "formaDePagoID": 5,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 5, "descripcion": "Transferencia Bancaria"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "6:SALICE")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-8"),
    )

    payments, _meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].empresa == "ALARCON"


def test_fetch_receipts_and_payments_gesi_accepts_numero_without_nro_recibo_pm(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 6,
                "numero": 1,
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "vendedorID": 1,
                "formaDePagoID": 9,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 9, "descripcion": "Transferencia Bancaria"},
        ],
    }

    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-9"),
    )

    payments, meta = fetch_receipts_and_payments(60, None)
    assert len(payments) == 1
    assert payments[0].nro_recibo == "1"
    assert meta["payments_count"] == 1


def test_fetch_receipts_and_payments_prefers_sucursal_id_for_empresa(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 3,
                "sucursal_id": 6,
                "numero": 1,
                "codigoDeImportacion": "PMCBR_70123",
                "clienteID": 33119,
                "razonSocial": "Cliente",
                "formaDePagoID": 9,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 9, "descripcion": "Transferencia Bancaria"},
        ],
    }

    monkeypatch.setenv("RECEIPTS_API_EMPRESA_MAP", "3:SALICE,6:SALICE")
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-gesi-suc"),
    )

    payments, _meta = fetch_receipts_and_payments(15, None)
    assert len(payments) == 1
    assert payments[0].empresa == "ALARCON"


def test_fetch_receipts_and_payments_skips_non_pm_codigo_importacion(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 3,
                "codigoDeImportacion": "CBR_99999",
                "clienteID": 33119,
                "formaDePagoID": 9,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 9, "descripcion": "Transferencia Bancaria"},
        ],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-no-pm"),
    )

    payments, meta = fetch_receipts_and_payments(15, None)
    assert payments == []
    assert meta["payments_count"] == 0
    assert any("sin nro de recibo utilizable" in w for w in meta.get("external_warnings", []))


def test_fetch_receipts_and_payments_accepts_numeric_codigo_importacion(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 3,
                "codigoDeImportacion": "70111",
                "clienteID": 33119,
                "formaDePagoID": 9,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 9, "descripcion": "Transferencia Bancaria"},
        ],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-pm-num"),
    )

    payments, meta = fetch_receipts_and_payments(15, None)
    assert len(payments) == 1
    assert payments[0].nro_recibo == "70111"
    assert meta["payments_count"] == 1


def test_fetch_receipts_and_payments_keeps_non_bankable_efectivo_for_later_stages(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 3,
                "codigoDeImportacion": "PMCBR_70112",
                "clienteID": 33119,
                "formaDePagoID": 8,
                "formaDePago": "EFECTIVO",
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"formaDePagoID": 8, "descripcion": "EFECTIVO"},
        ],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-efec"),
    )

    payments, meta = fetch_receipts_and_payments(15, None)
    assert len(payments) == 1
    assert payments[0].medio_pago == "EFECTIVO"
    assert meta["payments_count"] == 1
    assert meta["medio_bancarizable_stats"]["NON_BANKABLE"] == 1


def test_fetch_receipts_and_payments_maps_formas_pago_snake_case(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 3,
                "codigoDeImportacion": "PMCBR_70120",
                "clienteID": 33119,
                "forma_pago_id": 9,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [
            {"forma_pago_id": 9, "descripcion": "Transferencia Bancaria"},
        ],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-forma-snake"),
    )

    payments, meta = fetch_receipts_and_payments(15, None)
    assert len(payments) == 1
    assert payments[0].medio_pago == "Transferencia Bancaria"
    assert meta["medio_bancarizable_stats"]["BANKABLE"] == 1


def test_fetch_receipts_and_payments_unknown_medio_is_kept_as_sin_medio(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 3,
                "codigoDeImportacion": "PMCBR_70121",
                "clienteID": 33119,
                "formaDePagoID": 999,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
            }
        ],
        "formasDePago": [],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-forma-unknown"),
    )

    payments, meta = fetch_receipts_and_payments(15, None)
    assert len(payments) == 1
    assert payments[0].medio_pago == "SIN_MEDIO_API"
    assert meta["payments_count"] == 1
    assert meta["medio_bancarizable_stats"]["UNKNOWN"] == 1


def test_fetch_receipts_and_payments_uses_medios_de_pago_by_valor_id(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 3,
                "sucursalID": 3,
                "codigoDeImportacion": "PMCBR_70122",
                "clienteID": 33119,
                "formaDePagoID": 0,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
                "detalleDeValores": [
                    {"valorID": 2, "importe": 1234.0, "tipoMovimiento": "E"},
                ],
            },
            {
                "empresaID": 3,
                "sucursalID": 3,
                "codigoDeImportacion": "PMCBR_70123",
                "clienteID": 33119,
                "formaDePagoID": 0,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 1234.0,
                "detalleDeValores": [
                    {"valorID": 1, "importe": 1234.0, "tipoMovimiento": "E"},
                ],
            },
        ],
        "formasDePago": [],
        "mediosDePago": [
            {"empresaID": 3, "valorID": 2, "descripcion": "Transf. Bancaria", "tipo": "B"},
            {"empresaID": 3, "valorID": 1, "descripcion": "Efectivo", "tipo": "E"},
        ],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-medios-valor"),
    )

    payments, meta = fetch_receipts_and_payments(15, None)
    assert len(payments) == 2
    by_recibo = {p.nro_recibo: p for p in payments}
    assert "Transf. Bancaria" in by_recibo["70122"].medio_pago
    assert "Efectivo" in by_recibo["70123"].medio_pago
    assert meta["medio_bancarizable_stats"]["BANKABLE"] >= 1
    assert meta["medio_bancarizable_stats"]["NON_BANKABLE"] >= 1


def test_fetch_receipts_and_payments_uses_bankable_subset_amount_from_detalle(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 3,
                "sucursalID": 3,
                "codigoDeImportacion": "PMCBR_69302",
                "clienteID": 33119,
                "formaDePagoID": 0,
                "fechaDeEmision": "2026-02-26",
                "importeTotal": 638658.35,
                "detalleDeValores": [
                    {"valorID": 2, "importe": 450000.0, "tipoMovimiento": "E"},
                    {"valorID": 1, "importe": 188658.35, "tipoMovimiento": "E"},
                ],
            }
        ],
        "formasDePago": [],
        "mediosDePago": [
            {"empresaID": 3, "valorID": 2, "descripcion": "Transf. Bancaria", "tipo": "B"},
            {"empresaID": 3, "valorID": 1, "descripcion": "Efectivo", "tipo": "E"},
        ],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-medios-mixed-amount"),
    )

    payments, _meta = fetch_receipts_and_payments(15, None)
    assert len(payments) == 1
    assert payments[0].nro_recibo == "69302"
    assert payments[0].importe_pago == 450000.0


def test_fetch_receipts_and_payments_prioritizes_medios_label_over_legacy_text(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 3,
                "sucursalID": 3,
                "codigoDeImportacion": "PMCBR_70124",
                "clienteID": 33119,
                "formaDePagoID": 0,
                "formaDePago": "NO_INFORMADO",
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 500.0,
                "detalleDeValores": [{"valorID": 2, "importe": 500.0, "tipoMovimiento": "E"}],
            }
        ],
        "formasDePago": [],
        "mediosDePago": [
            {"empresaID": 3, "valorID": 2, "descripcion": "Transf. Bancaria", "tipo": "B"},
        ],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-medios-label"),
    )

    payments, _meta = fetch_receipts_and_payments(15, None)
    assert len(payments) == 1
    assert payments[0].medio_pago == "Transf. Bancaria"


def test_fetch_receipts_and_payments_maps_float_like_ids(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 3,
                "sucursalID": "6.0",
                "codigoDeImportacion": "PMCBR_70130",
                "clienteID": 33119,
                "formaDePagoID": 0,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 500.0,
                "detalleDeValores": [{"valorID": "2.0", "importe": 500.0, "tipoMovimiento": "E"}],
            }
        ],
        "formasDePago": [],
        "mediosDePago": [
            {"empresaID": "6.0", "valorID": "2.0", "descripcion": "Transf. Bancaria", "tipo": "B"},
        ],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None: _Resp(payload, request_id="r-float-ids"),
    )

    payments, meta = fetch_receipts_and_payments(15, None)
    assert len(payments) == 1
    assert payments[0].empresa == "ALARCON"
    assert payments[0].medio_pago == "Transf. Bancaria"
    assert meta["medio_bancarizable_stats"]["BANKABLE"] == 1


def test_fetch_receipts_and_payments_gba_sets_business_empresa_to_gba(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 6,
                "sucursalID": 6,
                "codigoDeImportacion": "PMCBR_80001",
                "clienteID": 33119,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 501.0,
                "detalleDeValores": [],
            }
        ],
        "formasDePago": [],
        "mediosDePago": [],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None, start_date=None, end_date=None: _Resp(payload, request_id="r-gba-emp"),
    )

    payments, _meta = fetch_receipts_and_payments(15, "GBA")
    assert len(payments) == 1
    assert payments[0].empresa == "GBA"


def test_fetch_receipts_and_payments_gba_marks_unknown_without_detail(monkeypatch):
    payload = {
        "comprobantes": [
            {
                "empresaID": 2,
                "sucursalID": 5,
                "codigoDeImportacion": "PMCBR_80002",
                "clienteID": 33119,
                "fechaDeEmision": "2026-02-18",
                "importeTotal": 501.25,
                "detalleDeValores": [],
            }
        ],
        "formasDePago": [],
        "mediosDePago": [],
    }
    monkeypatch.setattr(
        "src.conciliador.external.service.fetch_receipts_payload",
        lambda days, empresa_filter=None, start_date=None, end_date=None: _Resp(payload, request_id="r-gba-odd"),
    )

    payments, meta = fetch_receipts_and_payments(15, "GBA")
    assert len(payments) == 1
    assert payments[0].medio_pago == "SIN_MEDIO_API"
    assert meta["medio_bancarizable_stats"]["UNKNOWN"] == 1
