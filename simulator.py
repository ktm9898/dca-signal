"""
DCA Signal - Quantitative Dollar-Cost Averaging Simulation Engine
Simulates multi-cycle capital allocation, phase-based entry rules, and take-profit mechanics.
"""

import os
import json
import math
from datetime import datetime
from typing import Dict, List, Any, Optional

class DCASimulator:
    def __init__(
        self,
        candles: List[Dict[str, Any]],
        initial_capital: float = 10000000.0,
        portions: int = 40,
        phase_split_round: int = 20,
        first_half_rule: str = "always",       # "always"
        second_half_rule: str = "below_threshold", # "below_threshold", "below_threshold_or_half", "always"
        second_half_threshold_pct: float = 0.0,    # 0.0 means close <= avg_cost, -5.0 means close <= avg_cost * 0.95
        target_profit_pct: float = 10.0,
        fee_pct: float = 0.1,                   # 0.1% per trade
        slippage_pct: float = 0.05,
        reinvest_profit: bool = False,          # True: seed increases by profits; False: fixed seed per cycle
        max_round_action: str = "hold",         # "hold": stop buying when seed exhausted; "fail_cut": stop-loss
        stop_loss_pct: Optional[float] = None,  # None or e.g. -30.0
        start_date: Optional[str] = None,
        end_date: Optional[str] = None
    ):
        self.candles = [
            c for c in candles 
            if (not start_date or c["date"] >= start_date) and (not end_date or c["date"] <= end_date)
        ]
        self.initial_capital = float(initial_capital)
        self.portions = int(portions)
        self.phase_split_round = int(phase_split_round)
        self.first_half_rule = first_half_rule
        self.second_half_rule = second_half_rule
        self.second_half_threshold_pct = float(second_half_threshold_pct)
        self.target_profit_pct = float(target_profit_pct)
        self.fee_pct = float(fee_pct) / 100.0
        self.slippage_pct = float(slippage_pct) / 100.0
        self.reinvest_profit = reinvest_profit
        self.max_round_action = max_round_action
        self.stop_loss_pct = stop_loss_pct

    def run(self) -> Dict[str, Any]:
        if not self.candles:
            return {"error": "No candles available for the specified range."}

        current_capital_pool = self.initial_capital
        cycle_base_capital = self.initial_capital

        # Active cycle state
        cycle_idx = 1
        cycle_round = 0          # How many portions purchased
        shares_held = 0.0
        total_invested = 0.0     # Cash spent in current cycle
        cycle_start_date = None
        cycle_start_idx = 0
        cycle_trades = []        # Daily trade actions in current cycle

        completed_cycles = []
        equity_curve = []
        
        # Benchmark Buy & Hold tracking
        bh_shares = self.initial_capital / (self.candles[0]["open"] * (1 + self.fee_pct + self.slippage_pct))
        
        realized_pnl_total = 0.0
        peak_equity = self.initial_capital
        max_drawdown_pct = 0.0
        max_cycle_days = 0

        for idx, bar in enumerate(self.candles):
            date = bar["date"]
            open_p = bar["open"]
            high_p = bar["high"]
            low_p = bar["low"]
            close_p = bar["close"]

            # Current average cost per share
            avg_cost = (total_invested / shares_held) if shares_held > 0 else 0.0

            # ----------------------------------------------------
            # 1. Check Take-Profit & Stop-Loss (Exit evaluation)
            # ----------------------------------------------------
            target_sell_price = avg_cost * (1.0 + self.target_profit_pct / 100.0) if avg_cost > 0 else 0.0
            is_take_profit = False
            is_stop_loss = False
            exit_price = 0.0

            if shares_held > 0:
                # Stop loss check
                if self.stop_loss_pct is not None and avg_cost > 0:
                    stop_price = avg_cost * (1.0 + self.stop_loss_pct / 100.0)
                    if low_p <= stop_price:
                        is_stop_loss = True
                        exit_price = min(open_p, stop_price) * (1.0 - self.slippage_pct)

                # Take profit check (if not stopped out)
                if not is_stop_loss and target_sell_price > 0:
                    if high_p >= target_sell_price:
                        is_take_profit = True
                        executed_p = max(open_p, target_sell_price) if open_p >= target_sell_price else target_sell_price
                        exit_price = executed_p * (1.0 - self.slippage_pct)

            if is_take_profit or is_stop_loss:
                # Process Full Exit
                gross_revenue = shares_held * exit_price
                fee = gross_revenue * self.fee_pct
                net_revenue = gross_revenue - fee
                net_profit = net_revenue - total_invested
                return_pct = (net_profit / total_invested) * 100.0 if total_invested > 0 else 0.0

                duration_days = idx - cycle_start_idx + 1
                max_cycle_days = max(max_cycle_days, duration_days)
                realized_pnl_total += net_profit

                cycle_record = {
                    "cycle_num": cycle_idx,
                    "status": "PROFIT" if is_take_profit else "STOP_LOSS",
                    "start_date": cycle_start_date,
                    "end_date": date,
                    "duration_days": duration_days,
                    "max_rounds_used": cycle_round,
                    "total_invested": round(total_invested, 2),
                    "gross_revenue": round(gross_revenue, 2),
                    "net_profit": round(net_profit, 2),
                    "return_pct": round(return_pct, 2),
                    "avg_cost": round(avg_cost, 4),
                    "exit_price": round(exit_price, 4),
                    "trades_count": len(cycle_trades)
                }
                completed_cycles.append(cycle_record)

                # Reinvestment or Fixed seed
                if self.reinvest_profit:
                    current_capital_pool += net_profit
                    cycle_base_capital = current_capital_pool

                # Reset state for next cycle
                cycle_idx += 1
                cycle_round = 0
                shares_held = 0.0
                total_invested = 0.0
                avg_cost = 0.0
                cycle_start_date = None
                cycle_trades = []

            # ----------------------------------------------------
            # 2. Check Daily Buying Conditions (Entry evaluation)
            # ----------------------------------------------------
            portion_size = cycle_base_capital / self.portions
            funds_left = cycle_base_capital - total_invested
            
            # Can we buy today?
            can_buy_more = (cycle_round < self.portions) or (self.max_round_action != "hold")
            
            if can_buy_more and funds_left >= (portion_size * 0.1):
                # Determine current phase: 1st half vs 2nd half
                is_first_half = (cycle_round < self.phase_split_round)
                
                budget_today = 0.0
                order_fill_price = close_p * (1.0 + self.slippage_pct)

                if is_first_half:
                    # First half: regular fixed amount DCA
                    budget_today = min(portion_size, funds_left)
                else:
                    # Second half: defensive threshold rule
                    target_threshold_price = avg_cost * (1.0 + self.second_half_threshold_pct / 100.0)
                    is_below_threshold = (avg_cost == 0 or close_p <= target_threshold_price)

                    if self.second_half_rule == "below_threshold":
                        # Buy only if today's close is <= threshold (e.g. <= avg_cost or <= avg_cost - X%)
                        if is_below_threshold:
                            budget_today = min(portion_size, funds_left)
                        else:
                            budget_today = 0.0  # Pass, save cash
                    elif self.second_half_rule == "below_threshold_or_half":
                        if is_below_threshold:
                            budget_today = min(portion_size, funds_left)
                        else:
                            budget_today = min(portion_size * 0.5, funds_left)
                    elif self.second_half_rule == "always":
                        budget_today = min(portion_size, funds_left)

                if budget_today > 0 and order_fill_price > 0:
                    bought_shares = (budget_today * (1.0 - self.fee_pct)) / order_fill_price
                    if bought_shares > 0:
                        if cycle_start_date is None:
                            cycle_start_date = date
                            cycle_start_idx = idx

                        shares_held += bought_shares
                        total_invested += budget_today
                        cycle_round += 1
                        avg_cost = total_invested / shares_held

                        cycle_trades.append({
                            "date": date,
                            "round": cycle_round,
                            "price": round(order_fill_price, 4),
                            "shares": round(bought_shares, 4),
                            "cost": round(budget_today, 2),
                            "avg_cost": round(avg_cost, 4)
                        })

            # ----------------------------------------------------
            # 3. Track Daily Mark-to-Market Portfolio Value & MDD
            # ----------------------------------------------------
            current_shares_val = shares_held * close_p
            cash_in_hand = (cycle_base_capital - total_invested) + (current_capital_pool - cycle_base_capital) + realized_pnl_total
            total_equity = cash_in_hand + current_shares_val

            # Benchmark Buy & Hold equity
            bh_equity = bh_shares * close_p

            if total_equity > peak_equity:
                peak_equity = total_equity
            
            dd_pct = ((total_equity - peak_equity) / peak_equity) * 100.0 if peak_equity > 0 else 0.0
            if dd_pct < max_drawdown_pct:
                max_drawdown_pct = dd_pct

            equity_curve.append({
                "date": date,
                "equity": round(total_equity, 2),
                "cash": round(cash_in_hand, 2),
                "position_value": round(current_shares_val, 2),
                "cycle_num": cycle_idx,
                "cycle_round": cycle_round,
                "shares_held": round(shares_held, 4),
                "avg_cost": round(avg_cost, 4),
                "close": round(close_p, 4),
                "drawdown_pct": round(dd_pct, 2),
                "bh_equity": round(bh_equity, 2)
            })

        # Calculate final metrics
        final_equity = equity_curve[-1]["equity"]
        total_return_pct = ((final_equity - self.initial_capital) / self.initial_capital) * 100.0
        
        # Calculate CAGR
        total_days = (datetime.strptime(self.candles[-1]["date"], "%Y-%m-%d") - datetime.strptime(self.candles[0]["date"], "%Y-%m-%d")).days
        years = max(total_days / 365.25, 0.1)
        cagr_pct = ((final_equity / self.initial_capital) ** (1.0 / years) - 1.0) * 100.0

        # Benchmark CAGR
        bh_final = equity_curve[-1]["bh_equity"]
        bh_return_pct = ((bh_final - self.initial_capital) / self.initial_capital) * 100.0
        bh_cagr_pct = ((bh_final / self.initial_capital) ** (1.0 / years) - 1.0) * 100.0

        # Yearly breakdown
        yearly_stats = {}
        for c in completed_cycles:
            end_yr = c["end_date"][:4]
            if end_yr not in yearly_stats:
                yearly_stats[end_yr] = {
                    "year": end_yr,
                    "completed_cycles": 0,
                    "realized_profit": 0.0,
                    "max_rounds": 0
                }
            yearly_stats[end_yr]["completed_cycles"] += 1
            yearly_stats[end_yr]["realized_profit"] += c["net_profit"]
            yearly_stats[end_yr]["max_rounds"] = max(yearly_stats[end_yr]["max_rounds"], c["max_rounds_used"])

        # Also get yearly equity return and yearly MDD from equity curve
        curve_by_year = {}
        for eq in equity_curve:
            yr = eq["date"][:4]
            if yr not in curve_by_year:
                curve_by_year[yr] = []
            curve_by_year[yr].append(eq)

        yearly_breakdown = []
        for yr, pts in sorted(curve_by_year.items()):
            start_eq = pts[0]["equity"]
            end_eq = pts[-1]["equity"]
            yr_ret = ((end_eq - start_eq) / start_eq) * 100.0
            
            yr_peak = start_eq
            yr_mdd = 0.0
            for p in pts:
                if p["equity"] > yr_peak:
                    yr_peak = p["equity"]
                cur_dd = ((p["equity"] - yr_peak) / yr_peak) * 100.0
                if cur_dd < yr_mdd:
                    yr_mdd = cur_dd

            c_info = yearly_stats.get(yr, {"completed_cycles": 0, "realized_profit": 0.0, "max_rounds": 0})
            yearly_breakdown.append({
                "year": yr,
                "start_equity": round(start_eq, 2),
                "end_equity": round(end_eq, 2),
                "return_pct": round(yr_ret, 2),
                "completed_cycles": c_info["completed_cycles"],
                "realized_profit": round(c_info["realized_profit"], 2),
                "max_rounds_used": c_info["max_rounds"],
                "mdd_pct": round(yr_mdd, 2)
            })

        avg_cycle_days = (
            sum(c["duration_days"] for c in completed_cycles) / len(completed_cycles)
            if completed_cycles else 0.0
        )

        return {
            "summary": {
                "initial_capital": self.initial_capital,
                "final_equity": round(final_equity, 2),
                "total_profit": round(final_equity - self.initial_capital, 2),
                "total_return_pct": round(total_return_pct, 2),
                "cagr_pct": round(cagr_pct, 2),
                "max_drawdown_pct": round(max_drawdown_pct, 2),
                "completed_cycles_count": len(completed_cycles),
                "avg_cycle_days": round(avg_cycle_days, 1),
                "max_cycle_days": max_cycle_days,
                "years": round(years, 2),
                "benchmark_return_pct": round(bh_return_pct, 2),
                "benchmark_cagr_pct": round(bh_cagr_pct, 2),
                "in_progress_cycle": {
                    "cycle_num": cycle_idx,
                    "round": cycle_round,
                    "shares": round(shares_held, 4),
                    "invested": round(total_invested, 2),
                    "avg_cost": round(avg_cost, 4),
                    "unrealized_pnl_pct": round(((self.candles[-1]["close"] - avg_cost) / avg_cost) * 100.0, 2) if avg_cost > 0 else 0.0,
                    "start_date": cycle_start_date
                }
            },
            "parameters": {
                "portions": self.portions,
                "phase_split_round": self.phase_split_round,
                "second_half_rule": self.second_half_rule,
                "second_half_threshold_pct": self.second_half_threshold_pct,
                "target_profit_pct": self.target_profit_pct,
                "reinvest_profit": self.reinvest_profit
            },
            "yearly_breakdown": yearly_breakdown,
            "completed_cycles": completed_cycles,
            "equity_curve": equity_curve[::max(1, len(equity_curve)//500)]
        }

if __name__ == "__main__":
    data_file = os.path.join(os.path.dirname(__file__), "data", "TQQQ.json")
    if os.path.exists(data_file):
        with open(data_file, "r", encoding="utf-8") as f:
            tqqq_data = json.load(f)
        
        sim = DCASimulator(
            candles=tqqq_data["candles"],
            initial_capital=10000000,
            portions=40,
            phase_split_round=20,
            second_half_rule="below_threshold",
            second_half_threshold_pct=0.0,
            target_profit_pct=10.0
        )
        res = sim.run()
        print("=== DCA Simulator (TQQQ 10Y) Test Result ===")
        print(f"Final Equity: {int(res['summary']['final_equity']):,}원 ({res['summary']['total_return_pct']}%)")
        print(f"CAGR: {res['summary']['cagr_pct']}% | MDD: {res['summary']['max_drawdown_pct']}%")
        print(f"Completed Cycles: {res['summary']['completed_cycles_count']}")
