# ALD automation toolkit

Tools that take retyping out of the quote and survey workflow. They sit upstream of the
existing Claude skills (`leak-detection-repair-quote`, `leak-detection-report`), which
already build the customer PDFs.

## Cost Engine generator

One master rate card (`rate_card.yaml`) and one small job file per quote. The output is the
Kawasaki-style Cost Engine: Rate Card sheet with a basis on every input, Quote Builder sheet
driven by formulas end to end, gross margin pricing (Sell = Cost / (1 - margin)).

```
python -m ald.cost_engine jobs/kawasaki_maryville.yaml          # writes output/ALD Quote - <no> - Cost Engine.xlsx
python tools/recalc.py "output/ALD Quote - <no> - Cost Engine.xlsx"   # optional: cache values
```

New job: copy `jobs/kawasaki_maryville.yaml`, change the quote header, compressor list and
any overrides. An override without a `basis:` is rejected. The rate card default for
`energy_rate` is a placeholder, and the workbook flags it until the job sets the plant rate.

Changing a rate (tech day, hotel, margin) in `rate_card.yaml` updates every Cost Engine
generated after that. Nothing is retyped per customer.

The Kawasaki job reproduces the hand-built ALD-2026-1013-KM workbook to the dollar
(survey $30,750, repair add-on $81,250, guaranteed savings $81,075). `tests/test_cost_engine.py`
checks that.

## Leak sheet template

`templates/ALD Leak Sheet Template.xlsx`, built by `python -m ald.leak_sheet`. Same column
layout as `Template Leak Sheet.xlsx`, so `leak-detection-report` reads it unchanged.

| Problem in the 2023-2025 templates | Fix |
|---|---|
| Pressure other than 150/125/100/75/50/25/10 psi breaks the row | Any 10-150 psi works; between curves CFM is linearly interpolated and the row is marked INTERPOLATED |
| VLOOKUP truncates fractional dB (42.5 dB is sized as 42 dB) | CFM computed from the UE Systems curve at the actual dB |
| % repaired uses COUNTA, so an "N" counts as repaired | Counts "Y" only |
| kWh for nitrogen, argon etc. is gas dollars divided by $/kWh | kWh shown for Air and Open Blow only |
| $0.06/kWh default, sheet named "2023" | Rate and year are labeled inputs; $/kWh placeholder is flagged until set |
| FILTER() in the Report sheet | SUMIFS, works in every Excel version |
| Open blows mixed into leak totals | "Open Blow" gas type, excluded from leak totals, reported separately |

ASSUMPTION: interpolating CFM linearly between two UE Systems pressure curves is ALD's
choice, not a UE Systems method. Rows that use it are labeled so they can be reviewed.

Gas costs for CO2, argon, helium, hydrogen and nitrogen are carried over from the old
templates. Their source is not documented, so the sheet says to confirm them against a gas
supplier invoice.

## Tests

```
pip install pytest
python -m pytest tests -q
```

The tests fill the templates with leaks, recalculate them in LibreOffice and check every
CFM, cost and kWh value against an independent Python calculation and against the
`leak-detection-report` reader.
