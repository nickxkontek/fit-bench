# FitBench

Least-squares curve fitting that runs in your browser or as an installed Windows app.

* Models: polynomial (any degree), proportional, logarithmic, exponential (with or without offset), power law, Gaussian peak, sinusoid, damped oscillation, logistic, saturation (Michaelis–Menten), and custom formulas
* Weighted fitting with y errors, x errors through effective variance, and a robust (Huber) mode for outliers
* Parameter uncertainties, χ², reduced χ², p-value, R², AICc, a 95% confidence band, and a model comparison table
* Import CSV, TSV, TXT and Excel (.xlsx) files
* Export to Excel (with live formulas), CSV, a text report, JSON, or a PNG of the plot
* Works offline once opened

## Files

|File|What it is|
|-|-|
|`index.html`|The app|
|`version.js`|The version number. **Raise it every time you upload an update.**|
|`sw.js`|Offline support and update delivery (no need to edit)|
|`xlsx-lite.js`|Excel import and export|
|`manifest.webmanifest`|App name, icons and file types for Windows and the Microsoft Store|
|`privacy.html`|Privacy policy (linked from the Store listing)|
|`icons/`, `screenshots/`|App icons and Store screenshots|



