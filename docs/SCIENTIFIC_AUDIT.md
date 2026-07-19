# Scientific Audit

## 1. Fatigue indices

### Simple IGFA

Formula:

IGFA = [RPE + MS + (10 - PF) + (10 - MF) + (10 - RC)] / 5

The index combines five subjective variables on a common 0–10 scale.

Status:

- Operational composite index developed for this software.
- Not presented as an externally validated diagnostic instrument.
- Higher values represent a less favourable subjective state.

### Weighted IGFA

Formula:

IGFA = 0.30·RPE + 0.25·MS + 0.20·(10-PF) +
       0.15·(10-MF) + 0.10·(10-RC)

Status:

- Operational weighting scheme.
- Greater weight is assigned to perceived exertion and muscle soreness.
- The weights must be described as predefined analytical criteria unless
  external validation is performed.

## 2. Classification thresholds

Current thresholds:

- < 2: very low fatigue
- 2–<4: low fatigue
- 4–<6: moderate fatigue
- 6–<8: high fatigue
- >= 8: very high fatigue

These thresholds are descriptive categories used by the application.

They must not be described as validated clinical or performance cut-offs
without a specific validation study.

## 3. PRE–POST interpretation

The software calculates:

Delta = POST - PRE

Interpretation:

- RPE, MS and IGFA:
  positive delta = less favourable response
  negative delta = more favourable response

- PF, MF and RC:
  positive delta = more favourable response
  negative delta = less favourable response

This interpretation is internally coherent and must remain explicit in every
table, graph and report.

## 4. Linear trends

The application estimates a linear trend using the sequential position of
the observations.

Required correction:

- Sort matches chronologically by date and time before estimating trends.
- Do not interpret the slope as a causal effect.
- Suppress interpretation when fewer than three observations are available.
- Treat R² below 0.10 as insufficient for meaningful descriptive reporting.

## 5. Competitive difficulty

Competitive difficulty categories are ordinal.

Required rule:

- Compare observed means by category.
- Do not interpret a linear regression across category labels as a temporal
  or causal trend.
- Report sample size and standard deviation for each category.

## 6. Automatic findings

Automatic findings must use cautious language.

Recommended expressions:

- "The data are compatible with..."
- "A descriptive increase was observed..."
- "The available observations suggest..."
- "No inferential conclusion is made."

Avoid:

- "The match caused fatigue."
- "The referee deteriorated because of..."
- "A significant effect exists."

## 7. Minimum reporting requirements

Every generated report should state:

- number of valid observations;
- variables included;
- missing-data handling;
- formula used for each index;
- meaning of positive and negative deltas;
- descriptive nature of the analysis;
- absence of automatic inferential conclusions.

## 8. Priority corrections

1. Order matches by date and time before trend analysis.
2. Remove trend regression from competitive-category graphs.
3. Label IGFA thresholds as operational.
4. Require at least three observations for trend interpretation.
5. Add missing-data information to the report.
6. Preserve cautious scientific language.
