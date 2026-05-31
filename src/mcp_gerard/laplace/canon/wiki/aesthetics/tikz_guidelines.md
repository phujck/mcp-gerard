# TikZ Guidelines & Visual Aesthetics

All visual outputs must meet the "Nature-Style" benchmark - see [figure_standard](canon://skills/figure_standard) and [Visual Aesthetics](canon://aesthetics/visual_mechanics.md). Tolerances for visual flaws are zero.

## Axioms of Visual Proof
* **The Interesting Regime**: All plots must maximize the view of the "interesting/dynamic regime" while proving steady-state persistence. Do not waste space on dead zones.
* **TeX-Native Elements**: Axes, labels, and legends must be formatted natively in TeX using `pgfplots`. No imported PNG/PDF text. Enforce with [tikz_mechanic](canon://skills/tikz_mechanic).
* **Bounding Box Sanctity**: Absolute prohibition on text overlaying drawn objects. Elements must respect bounding boxes and font scaling.
* **Graphical Mapping**: Schematics must not be generic; they must contain graphical representations mapping directly to every core piece of the governing equations.
