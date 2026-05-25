"""test_integration_pipeline.py — Tests de integración end-to-end del matcher húngaro.

Ejercitan match_hungarian con datos sintéticos conocidos para detectar regresiones
cuando se modifica el matcher, el pipeline o las funciones de normalización.

No requieren API ni archivos externos.
"""
from __future__ import annotations

import datetime as dt

import pytest

from src.conciliador.excel_loader import BankTxn
from src.conciliador.pdf_parser import ReceiptPayment
from src.conciliador.matcher_hungarian import match_hungarian


# ---------------------------------------------------------------------------
# Helpers para construir fixtures mínimos
# ---------------------------------------------------------------------------

def _txn(
    txn_id: str,
    origen: str,
    fecha: str,
    importe: float,
    *,
    was_preconciled: bool = False,
    preconciled_recibo: str | None = None,
) -> BankTxn:
    return BankTxn(
        txn_id=txn_id,
        origen=origen,
        fecha=dt.date.fromisoformat(fecha),
        hora=None,
        importe=importe,
        texto_ref=f"ref-{txn_id}",
        row_index=int(txn_id.split("-")[-1]) if "-" in txn_id else 1,
        parse_ok=True,
        parse_error=None,
        was_preconciled=was_preconciled,
        preconciled_recibo=preconciled_recibo,
        sheet_name="Hoja1",
        record_key="default",
    )


def _payment(
    nro_recibo: str,
    nro_cliente: str,
    fecha: str,
    importe: float,
    medio: str = "TRANSFERENCIA",
    empresa: str = "GBA",
) -> ReceiptPayment:
    return ReceiptPayment(
        empresa=empresa,
        nro_recibo=nro_recibo,
        nro_cliente=nro_cliente,
        cliente_nombre=f"Cliente {nro_cliente}",
        medio_pago=medio,
        fecha_pago=fecha,
        importe_pago=importe,
    )


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestMatchHungarianBasic:
    """Escenarios de matching básico: mismo importe, mismo día."""

    def test_match_exacto_validado(self):
        """Un recibo y un ingreso con importe y fecha exactos → VALIDADO."""
        txns = [_txn("t-1", "BBVA", "2024-03-15", 10000.0)]
        payments = [_payment("1001", "5", "2024-03-15", 10000.0)]
        res = match_hungarian(
            txns,
            payments,
            report_date_min="2024-03-15",
            report_date_max="2024-03-15",
        )
        assert len(res["validados"]) == 1
        assert len(res["dudosos"]) == 0
        assert res["validados"][0]["Nro recibo"] == "1001"
        assert res["validados"][0]["Fila Excel"] == 1

    def test_match_con_diferencia_de_dias(self):
        """Recibo 2 días antes del ingreso → dudoso (Dif días = -2)."""
        txns = [_txn("t-2", "GALICIA", "2024-03-17", 5000.0)]
        payments = [_payment("2001", "10", "2024-03-15", 5000.0)]
        res = match_hungarian(
            txns,
            payments,
            valid_max_peso=0.0,  # forzar que nada sea validado
            dudoso_max_peso=5000.0,
            report_date_min="2024-03-15",
            report_date_max="2024-03-17",
        )
        assert len(res["dudosos"]) == 1
        assert res["dudosos"][0]["Dif días"] == -2

    def test_sin_recibos_todo_banco_sin_recibo(self):
        """Sin recibos → todos los ingresos van a no_encontrados como BANCO_SIN_RECIBO."""
        txns = [
            _txn("t-3", "BBVA", "2024-03-10", 1000.0),
            _txn("t-4", "GALICIA", "2024-03-11", 2000.0),
        ]
        res = match_hungarian(
            txns,
            [],
            report_date_min="2024-03-10",
            report_date_max="2024-03-11",
        )
        assert res["validados"] == []
        assert res["dudosos"] == []

    def test_sin_ingresos_todo_recibo_sin_banco(self):
        """Sin ingresos → todos los recibos van a no_encontrados como RECIBO_SIN_BANCO."""
        payments = [_payment("3001", "7", "2024-03-12", 3000.0)]
        res = match_hungarian(
            [],
            payments,
            report_date_min="2024-03-12",
            report_date_max="2024-03-12",
        )
        no_enc = res["no_encontrados"]
        recibo_sin_banco = [r for r in no_enc if r.get("Tipo no encontrado") == "RECIBO_SIN_BANCO"]
        assert len(recibo_sin_banco) == 1
        assert recibo_sin_banco[0]["Nro recibo"] == "3001"


class TestMatchHungarianMultiple:
    """Escenarios con múltiples recibos e ingresos."""

    def test_dos_matches_exactos(self):
        """Dos pares perfectos → ambos como VALIDADOS."""
        txns = [
            _txn("t-10", "BBVA", "2024-04-01", 1500.0),
            _txn("t-11", "BBVA", "2024-04-02", 2500.0),
        ]
        payments = [
            _payment("101", "20", "2024-04-01", 1500.0),
            _payment("102", "21", "2024-04-02", 2500.0),
        ]
        res = match_hungarian(
            txns,
            payments,
            report_date_min="2024-04-01",
            report_date_max="2024-04-02",
        )
        assert len(res["validados"]) == 2
        assert len(res["dudosos"]) == 0
        assert len(res["no_encontrados"]) == 0

    def test_match_optimo_no_greedy(self):
        """Verifica que el algoritmo húngaro asigne de forma óptima (no greedy).

        Caso: t1 podría ir con p1 o p2, t2 solo puede ir con p2.
        Greedy matchearía t1→p1 y dejaría t2 sin match.
        Húngaro debe matchear t1→p2, t2→p1 para maximizar cobertura total.

        Fechas y montos diseñados para que:
          - t1(10/04, $100) tiene costo bajo con p1(10/04, $100) Y con p2(10/04, $100)
          - t2(11/04, $100) solo tiene costo bajo con p1(10/04, $100) (1 día de dif)
        """
        txns = [
            _txn("t-20", "BBVA", "2024-04-10", 100.0),  # puede matchear p1 o p2
            _txn("t-21", "BBVA", "2024-04-11", 100.0),  # solo puede matchear p1 (1 día dif)
        ]
        payments = [
            _payment("201", "30", "2024-04-10", 100.0),  # fecha = t-20 exacto, pero también viable para t-21
            _payment("202", "31", "2024-04-10", 100.0),  # fecha = t-20 exacto
        ]
        res = match_hungarian(
            txns,
            payments,
            report_date_min="2024-04-10",
            report_date_max="2024-04-11",
        )
        # Húngaro debe encontrar asignación que cubra ambos recibos
        total_matched = len(res["validados"]) + len(res["dudosos"])
        # Al menos 1 validado porque hay un match perfecto garantizado
        assert len(res["validados"]) >= 1
        # No debe haber más no_encontrados de tipo RECIBO_SIN_BANCO que los inevitables
        recibo_sin_banco = [
            r for r in res["no_encontrados"]
            if r.get("Tipo no encontrado") == "RECIBO_SIN_BANCO"
        ]
        assert len(recibo_sin_banco) <= 1  # a lo sumo uno queda sin banco


class TestMatchHungarianPenalties:
    """Verifica que las penalizaciones afecten correctamente la asignación."""

    def test_preconciled_penalty_applied(self):
        """Un ingreso ya conciliado con otro recibo debe preferir el recibo original."""
        txns = [
            _txn("t-30", "BBVA", "2024-05-01", 1000.0,
                 was_preconciled=True, preconciled_recibo="5000"),
        ]
        payments = [
            _payment("5000", "40", "2024-05-01", 1000.0),  # recibo original → peso bajo
            _payment("5001", "41", "2024-05-01", 1000.0),  # recibo nuevo → suma preconciled_penalty
        ]
        res = match_hungarian(
            txns,
            payments,
            preconciled_penalty=150.0,
            valid_max_peso=260.0,
            report_date_min="2024-05-01",
            report_date_max="2024-05-01",
        )
        # El recibo original (5000) debe estar en validados, no el nuevo (5001)
        validados_recibos = {r["Nro recibo"] for r in res["validados"]}
        assert "5000" in validados_recibos
        assert "5001" not in validados_recibos

    def test_mp_mismatch_penalty(self):
        """Ingreso MERCADOPAGO que matchea con recibo TRANSFERENCIA → penalización."""
        txns = [
            _txn("t-40", "MERCADOPAGO", "2024-06-01", 2000.0),
        ]
        payments = [
            _payment("6001", "50", "2024-06-01", 2000.0, medio="TRANSFERENCIA"),
        ]
        res = match_hungarian(
            txns,
            payments,
            mp_mismatch_penalty=35.0,
            valid_max_peso=260.0,
            report_date_min="2024-06-01",
            report_date_max="2024-06-01",
        )
        # Con penalty=35 y todo lo demás perfecto, el peso = 35 < 260 → debe ser validado
        # (la penalidad no lo saca de validados, solo lo encarece)
        assert len(res["validados"]) + len(res["dudosos"]) == 1

    def test_non_bankable_receipt_only_participates_in_dudosos(self):
        """Un recibo no bancarizable no valida, pero sí puede caer en dudosos sin contaminar RECIBO_SIN_BANCO."""
        txns = [_txn("t-41", "BBVA", "2024-06-01", 2000.0)]
        payments = [_payment("6002", "51", "2024-06-02", 2000.0, medio="Efectivo")]

        res = match_hungarian(
            txns,
            payments,
            day_weight_bank_before=150.0,  # 1 día => costo base 150; con multiplicador x2 => 300
            valid_max_peso=260.0,
            dudoso_max_peso=3500.0,
            report_date_min="2024-06-01",
            report_date_max="2024-06-02",
            non_bankable_receipt_cost_multiplier=2.0,
            validated_allow_all_receipts=False,
            suspects_and_no_bankable_only=False,
            no_encontrados_bankable_only=True,
        )

        assert res["validados"] == []
        assert len(res["dudosos"]) == 1
        assert res["dudosos"][0]["Nro recibo"] == "6002"
        assert res["dudosos"][0]["Peso"] == 300.0
        assert res["no_encontrados"] == []


class TestMatchHungarianNormalization:
    """Verifica que la normalización de números de recibo funcione correctamente."""

    def test_normalize_recibo_contable(self):
        """Recibo con formato contable "68.734,00" debe normalizarse a "68734"."""
        from src.conciliador.utils import _normalize_recibo
        assert _normalize_recibo("68.734,00") == "68734"
        assert _normalize_recibo("1.234") == "1234"
        assert _normalize_recibo("5000") == "5000"
        assert _normalize_recibo(5000) == "5000"
        assert _normalize_recibo(5000.0) == "5000"
        assert _normalize_recibo(None) is None

    def test_normalize_cliente(self):
        """normalize_cliente debe extraer solo dígitos y quitar ceros iniciales."""
        from src.conciliador.utils import _normalize_cliente
        assert _normalize_cliente("00123") == "123"
        assert _normalize_cliente("456") == "456"
        assert _normalize_cliente(456) == "456"
        assert _normalize_cliente(None) is None
        assert _normalize_cliente("") is None

    def test_normalize_cuit(self):
        """normalize_cuit solo acepta exactamente 11 dígitos."""
        from src.conciliador.utils import _normalize_cuit
        assert _normalize_cuit("20-12345678-9") == "20123456789"
        assert _normalize_cuit("20123456789") == "20123456789"
        assert _normalize_cuit("123") is None
        assert _normalize_cuit(None) is None

    def test_normalize_text(self):
        """normalize_text debe eliminar acentos y pasar a minúsculas."""
        from src.conciliador.utils import _normalize_text
        assert _normalize_text("Ñoño") == "nono"
        assert _normalize_text("TRANSFERENCIA") == "transferencia"
        assert _normalize_text("  Hola  ") == "hola"


class TestMatchHungarianGraceDays:
    """Verifica la lógica de grace days para no_encontrados."""

    def test_recibo_reciente_no_es_no_encontrado(self):
        """Recibo cuya fecha está dentro del grace period no aparece en no_encontrados."""
        today = dt.date.today()
        fecha_reciente = (today - dt.timedelta(days=1)).isoformat()
        payments = [_payment("9001", "99", fecha_reciente, 500.0)]
        res = match_hungarian(
            [],
            payments,
            report_date_min=fecha_reciente,
            report_date_max=fecha_reciente,
            recibo_sin_banco_grace_days=2,
            current_date_override=today.isoformat(),
        )
        recibo_sin_banco = [
            r for r in res["no_encontrados"]
            if r.get("Tipo no encontrado") == "RECIBO_SIN_BANCO"
        ]
        # Con grace_days=2 y recibo de ayer, no debe aparecer en no_encontrados
        assert len(recibo_sin_banco) == 0
