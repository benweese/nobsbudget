==================================================================
BUDGET REBUILD – NO BS MODE: PROJECT PLAN (v3.0 - Updated 2025-09-24)
==================================================================

**CORE BUDGETING PHILOSOPHY:**
This system operates on a strict CASH FLOW basis. It is designed to track and budget
the actual cash leaving the primary bank account. Individual credit card transactions
are NOT imported. Instead, the full credit card PAYMENT is treated as a single,
budgeted expense. This provides a true "envelope style" grip on monthly cash flow.

------------------------------------------------------------------
PHASE 1: CORE ENGINE – COMPLETE & AUDITED ✅
------------------------------------------------------------------
1. Create Monthly Sheet ✅ DONE
2. Bank Import (CSV-Based) ✅ DONE
3. Codebase Refactor & Documentation ✅ DONE
   • All .gs files commented for clarity.
   • Bimonthly recurring item bug FIXED.
   • Core logic validated.

------------------------------------------------------------------
PHASE 2: INSIGHTS & INTELLIGENCE – IN PROGRESS 🟡
------------------------------------------------------------------
4. “Am I Screwed?” Dashboard (Current Focus)
   • ✅ Calculate total monthly cash income & expenses. (DONE)
   • ▶️ Calculate daily spend limit ("burn rate"). (NEXT)
   • Forecasted end-of-month balance.
   • Red/yellow/green alert system based on cash balance.
   • Comparison of cash flow vs. previous month.

5. Vendor & Burn Rate Charts
   • Top cash expenses by vendor this month.
   • Weekly spending trends from the bank account.

------------------------------------------------------------------
PHASE 3: SUPPORT SYSTEMS
------------------------------------------------------------------
6. Debt Goal Tracker
   • Connect to 'Debt Tracker' sheet data.
   • Visualize debt snowball/avalanche progress.
   • Calculate timelines to payoff.

7. Smart Alerts & Reporting
   • Generate a weekly summary email.
   • Alert on large, unexpected cash withdrawals.

------------------------------------------------------------------
PHASE 4: MAINTENANCE & OPTIMIZATION
------------------------------------------------------------------
8. Vendor Cleanup Tool
9. Performance Tuning ✅ DONE

------------------------------------------------------------------
PHASE 5: FUTURE INTEGRATION (Optional)
------------------------------------------------------------------
10. Plaid Integration
    • Note: Would require logic to filter for primary bank account
      transactions ONLY, ignoring credit cards.

------------------------------------------------------------------
CODE STRUCTURE (Current & Documented)
------------------------------------------------------------------
• main.gs → Menu items and triggers.
• bank_import.gs → CSV transaction logic for CASH accounts.
• vendor_utils.gs → Vendor matching and renaming logic.
• recurring_populator.gs → Monthly sheet creation logic.
• dashboard.gs → Calculations and logic for the main dashboard.
• utils.gs → Shared constants and helper functions.
• readme.gs → This project plan.