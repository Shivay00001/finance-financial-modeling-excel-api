# Financial Modeling Excel API — DCF

**What it does:** Builds a real discounted-cash-flow valuation from inputs
(revenue, per-year growth rates, EBIT margin, tax, capex/NWC assumptions,
WACC, terminal growth, net debt, shares) and returns full year-by-year numbers
plus a downloadable, formatted `.xlsx` workbook (inputs, FCF schedule,
valuation summary).

**Math:** `FCF = NOPAT − capex − ΔNWC`, discounted at WACC; terminal value via
Gordon growth `TV = FCF_n·(1+g)/(r−g)`; `EV = PV(FCF) + PV(TV)`;
`value/share = (EV − net debt) / shares`. No AI, no stubs — plain finance math.

## Run

```bash
pip install -r requirements.txt
uvicorn main:app --port 8000
```

## API

- `POST /dcf` — JSON model + valuation
- `POST /dcf/xlsx` — same model as downloadable `.xlsx`

```bash
curl -X POST localhost:8000/dcf -H 'Content-Type: application/json' -d '{
  "initial_revenue": 100, "growth_rates": [0.10, 0.08, 0.05],
  "ebit_margin": 0.20, "tax_rate": 0.25, "discount_rate": 0.10,
  "terminal_growth_rate": 0.02, "shares_outstanding": 10}'
curl -X POST localhost:8000/dcf/xlsx -H 'Content-Type: application/json' \
  -d '{...}' -o dcf_model.xlsx
```

## Tests

```bash
pytest -q   # includes hand-verified DCF arithmetic + xlsx integrity checks
```
