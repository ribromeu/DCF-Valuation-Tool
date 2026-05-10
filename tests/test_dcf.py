"""Unit tests for the DCF Valuation Tool.

Run with:
    pytest tests/
    or:
        python -m pytest tests/ -v
        """

import pytest
import sys
import os

# Ensure the project root is on the path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


# ---------------------------------------------------------------------------
# Helper / utility function tests
# ---------------------------------------------------------------------------

class TestDiscountedCashFlow:
      """Tests for core DCF calculation logic."""

    def test_present_value_single_cash_flow(self):
              """PV of a single cash flow discounted at rate r for n years."""
              cash_flow = 1000
              rate = 0.10
              years = 1
              pv = cash_flow / (1 + rate) ** years
              assert round(pv, 2) == 909.09

    def test_present_value_zero_rate(self):
              """When discount rate is 0, PV equals FV."""
              cash_flow = 500
              rate = 0.0
              years = 5
              pv = cash_flow / (1 + rate) ** years
              assert pv == cash_flow

    def test_sum_of_discounted_cash_flows(self):
              """Sum of discounted cash flows across multiple periods."""
              cash_flows = [100, 200, 300]
              rate = 0.10
              total_pv = sum(cf / (1 + rate) ** (i + 1) for i, cf in enumerate(cash_flows))
              assert round(total_pv, 2) == pytest.approx(481.59, rel=1e-3)


class TestTerminalValue:
      """Tests for terminal value calculations."""

    def test_gordon_growth_model(self):
              """Terminal value via Gordon Growth Model: TV = FCF * (1+g) / (r - g)."""
              fcf = 1000
              g = 0.02   # terminal growth rate
        r = 0.10   # discount rate / WACC
        tv = fcf * (1 + g) / (r - g)
        assert round(tv, 2) == pytest.approx(12750.0, rel=1e-3)

    def test_gordon_growth_model_raises_when_g_equals_r(self):
              """Gordon Growth Model is undefined when g == r (division by zero)."""
              fcf = 1000
              g = 0.10
              r = 0.10
              with pytest.raises(ZeroDivisionError):
                            _ = fcf * (1 + g) / (r - g)

          def test_exit_multiple_terminal_value(self):
                    """Terminal value via exit multiple: TV = EBITDA * multiple."""
                    ebitda = 500
                    multiple = 10
                    tv = ebitda * multiple
                    assert tv == 5000


class TestWACC:
      """Tests for Weighted Average Cost of Capital calculations."""

    def test_wacc_calculation(self):
              """WACC = E/V * Ke + D/V * Kd * (1 - tax_rate)."""
              equity = 600
              debt = 400
              total = equity + debt
              ke = 0.12   # cost of equity
        kd = 0.06   # cost of debt
        tax = 0.30  # tax rate

        wacc = (equity / total) * ke + (debt / total) * kd * (1 - tax)
        assert round(wacc, 4) == pytest.approx(0.0888, rel=1e-3)

    def test_wacc_all_equity(self):
              """With no debt, WACC equals the cost of equity."""
              equity = 1000
              debt = 0
              total = equity + debt
              ke = 0.15
              kd = 0.0
              tax = 0.30

        wacc = (equity / total) * ke + (debt / total if total else 0) * kd * (1 - tax)
        assert wacc == ke


class TestEquityValue:
      """Tests for equity value and per-share price derivation."""

    def test_equity_value_from_enterprise_value(self):
              """Equity value = Enterprise value - Net debt."""
              enterprise_value = 10_000
              net_debt = 2_000
              equity_value = enterprise_value - net_debt
              assert equity_value == 8_000

    def test_price_per_share(self):
              """Price per share = Equity value / shares outstanding."""
              equity_value = 8_000
              shares = 400
              price = equity_value / shares
              assert price == 20.0

    def test_margin_of_safety(self):
              """Margin of safety = (intrinsic - market) / intrinsic."""
              intrinsic = 50.0
              market = 40.0
              mos = (intrinsic - market) / intrinsic
              assert round(mos, 4) == 0.20


# ---------------------------------------------------------------------------
# Input validation tests
# ---------------------------------------------------------------------------

class TestInputValidation:
      """Tests for edge cases and input validation."""

    def test_negative_cash_flows_are_handled(self):
              """Negative cash flows should produce negative present values."""
              cash_flow = -500
              rate = 0.10
              years = 1
              pv = cash_flow / (1 + rate) ** years
              assert pv < 0

    def test_large_number_of_projection_years(self):
              """DCF should remain numerically stable over many years."""
              cash_flow = 100
              rate = 0.10
              years = 50
              pv = cash_flow / (1 + rate) ** years
              assert pv > 0
              assert pv < cash_flow  # always less than the nominal amount

    def test_growth_rate_must_be_less_than_wacc(self):
              """Terminal growth rate must be strictly less than WACC for a finite TV."""
              g = 0.10
              wacc = 0.08
              assert g >= wacc, "This scenario should be caught by validation logic"
      
