

## Plan: Fix Movimentacoes flow blocking on missing vencimento

### Problem
The input file `CPA.xlsx` has headers: `Nome Comp, CNPJ/CPF, Descrição, Valor, Forma de Pagamento, Observações` -- there is **no vencimento column at all**. The current code (lines 556-559) treats missing vencimento as a hard blocker and skips every row with `continue`, resulting in 0 despesas and 55 errors.

The tax default logic (lines 718-731) that could infer vencimento from tax rules runs **after** the skip, so it never gets a chance to execute.

### Solution

Restructure the flow in `process-movimentacoes/index.ts` so that vencimento is **not a hard blocker** during initial extraction. Instead, defer vencimento resolution to the processing phase where tax defaults and recurrence logic can fill in missing dates.

**Step 1 -- Remove the hard skip on missing vencimento (lines 552-560)**

- Still extract `vencRaw` and attempt `normalizeDate`, but do NOT `continue` if it fails.
- Store whatever was parsed (even empty string) and log a warning instead of skipping.
- Move the error/skip decision to later, after tax defaults and recurrence have had a chance to fill in the date.

**Step 2 -- Move vencimento inference earlier (restructure lines 718-731)**

During the output-building loop (line 625+), change the order:
1. First, try the parsed vencimento from input.
2. If empty, check tax defaults (already implemented) and apply default day.
3. If still empty, check if recurrence was detected -- if so, use current month + day 1 or a sensible default.
4. If still empty after all inference, **then** use a fallback: first day of next month as a generic "provisioned" date, and log a warning (not a hard skip).
5. Only skip the row if there's truly no name AND no vencimento (truly unusable).

**Step 3 -- Improve logging**

- Log how many rows had vencimento from input vs inferred from tax defaults vs inferred from recurrence vs fallback.
- Include the inference method in the error report for transparency.

### Impact
- Rows with no vencimento column will no longer be blocked; they'll get dates from tax defaults, recurrence rules, or a fallback.
- Only truly unusable rows (no name) will be skipped.
- No changes to other edge functions or UI.

