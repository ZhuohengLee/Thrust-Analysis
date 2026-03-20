# Thrust Analysis

This repository contains a single desktop GUI tool for cyclic thrust-data analysis:

- `thrust_analysis.py`

It is designed for thrust or force signals stored in `.xlsx`, `.xls`, or `.csv` files. The tool automatically estimates sampling rate, removes baseline drift, filters the signal, detects cycles, and exports both plots and summary tables.

## What the Script Does

The script processes each selected file in this order:

1. Load time and force/thrust data.
2. Detect the sampling rate from the time column.
3. Estimate a drifting baseline from the lower envelope of the signal.
4. Subtract the baseline to create a debiased signal.
5. Apply a low-pass filter to suppress noise.
6. Use FFT to estimate the dominant cycle period.
7. Detect peaks and their preceding baseline points for each cycle.
8. Compute per-cycle and per-file thrust statistics.
9. Export plots and a batch summary table.

## Run

```bash
python thrust_analysis.py
```

## Requirements

Install the required Python packages:

```bash
pip install pandas numpy matplotlib scipy openpyxl
```

## Supported Input Files

Supported formats:

- `.xlsx`
- `.xls`
- `.csv`

The script expects:

- one time column
- one thrust or force column

Automatic column detection rules:

- time column:
  - the first column whose name contains `time`
  - otherwise the first column in the file
- force column:
  - the first column whose name contains `force` or `thrust`
  - otherwise the second column, if present

## Main GUI Parameters

### Input Files

- `Add Single File`
  - add one file to the analysis queue
- `Add Multiple Files`
  - add multiple files at once
- `Add Folders`
  - scan one or more folders and collect all supported data files
- `Clear List`
  - clear the queued files

### Sampling

- `Detected Rate`
  - the detected sampling frequency in Hz
  - calculated from the median time-step in the input file
  - if detection fails, the tool falls back to `100.0 Hz`

Meaning:

- this value controls filtering and FFT period estimation
- if the sampling rate is wrong, the detected period and cycle spacing may also be wrong

### Cycle Detection

- `Detected Period`
  - the dominant cycle period estimated from FFT
  - displayed in seconds

Meaning:

- if the detected period is `1.20 s`, the algorithm assumes one full thrust cycle is roughly 1.20 seconds

### Analysis Options

- `Auto-detect peak parameters`
  - automatically adjusts peak-detection sensitivity from signal quality
  - recommended for most datasets

- `Processed Signal (Recommended)`
  - uses the filtered, baseline-corrected signal to measure cycle peaks and baselines
  - better for noisy data and routine reporting

- `Original Signal`
  - uses the raw signal values at detected cycle locations
  - useful when you want raw measured amplitudes instead of filtered values

- `Low-pass Cutoff (Hz)`
  - the cutoff frequency of the low-pass filter applied after baseline removal

Practical meaning:

- lower cutoff:
  - smoother signal
  - stronger noise suppression
  - may remove fast detail
- higher cutoff:
  - keeps more detail
  - may retain more noise

Recommended default:

- `5.0 Hz`

## Exported Results

For each source folder, the tool creates a `thrust_analysis_results` directory containing:

- one SVG figure per analyzed file
- one `batch_summary.xlsx` file per source folder

### Exported Plot

Each file generates:

- `<file_name>_analysis.svg`

The SVG contains six panels:

1. `Original Data & Baseline`
   - raw force signal and estimated baseline curve
2. `Baseline Removed`
   - debiased signal, computed as `force_raw - baseline_curve`
3. `Low-Pass Filter`
   - debiased signal plus filtered signal
4. `Removed Noise`
   - the difference between debiased and filtered signals
5. `Cycle Detection`
   - filtered signal with detected peaks and baselines marked
6. `Thrust vs Drag`
   - bar chart of cycle peak values and baseline-side values

### Exported Summary Table

The batch summary includes these columns:

- `Folder`
- `Folder_Path`
- `File`
- `Duration(s)`
- `Cycles`
- `Forward`
- `Backward`
- `Mean_Thrust(N)`
- `Mean_Drag(N)`
- `Mean_Net(N)`
- `Std_Net(N)`
- `Mean_Impulse(N*s)`
- `Total_Impulse(N*s)`
- `Detected_Period(s)`

## What Each Result Means

Below is the full meaning of the exported results and how each quantity is computed.

### `Folder`

The immediate parent folder name of the source file.

Meaning:

- useful when many datasets from different experiment folders are analyzed together

### `Folder_Path`

The full source folder path of the analyzed file.

Meaning:

- useful when two folders share the same short name
- helps trace every row back to the original dataset location

### `File`

The source filename without extension.

Meaning:

- the dataset identifier used in the plot name and summary row

### `Duration(s)`

Computed from the time axis:

```text
Duration = t[-1]
```

In this script, the exported duration is the last time value after loading and cleaning the file.

Meaning:

- the analyzed recording length in seconds

### `Cycles`

Computed as:

```text
Cycles = len(cycle_results)
```

Meaning:

- the number of valid thrust cycles detected in the file

Important note:

- only cycles with positive net thrust are kept
- a candidate cycle is accepted only if:

```text
net_thrust = peak_value - baseline_value > 0
```

### `Forward`

Computed as:

```text
Forward = number of cycles whose impulse > 0
```

Meaning:

- the number of cycles whose integrated filtered-signal area is positive
- in most datasets this is a count of cycles producing positive net impulse contribution

### `Backward`

Computed as:

```text
Backward = Cycles - Forward
```

Meaning:

- the number of detected cycles whose integrated impulse is not positive

### `Mean_Thrust(N)`

Computed as:

```text
Mean_Thrust = mean(peak_value for each detected cycle)
```

Where:

- `peak_value` is the cycle peak force
- it comes from the processed signal or original signal depending on the selected measurement mode

Meaning:

- the average peak force level across all detected cycles

### `Mean_Drag(N)`

Computed as:

```text
Mean_Drag = mean(baseline_value for each detected cycle)
```

Where:

- `baseline_value` is the valley or pre-peak reference value chosen for that cycle

Meaning:

- the average baseline-side force level before each thrust peak
- depending on the experiment, this may represent drag-side force, preload, or the low-force phase of a cycle

### `Mean_Net(N)`

Computed in two steps.

First, for each cycle:

```text
net_thrust = peak_value - baseline_value
```

Then:

```text
Mean_Net = mean(net_thrust for all detected cycles)
```

Meaning:

- the average net thrust amplitude per cycle
- this is often the most intuitive measure of how strong the cycle is after subtracting the baseline side

### `Std_Net(N)`

Computed as:

```text
Std_Net = std(net_thrust for all detected cycles)
```

Meaning:

- cycle-to-cycle variation of net thrust
- lower values mean the motion is more repeatable
- higher values suggest inconsistent cycle strength or unstable measurements

### `Mean_Impulse(N*s)`

For each cycle, the script computes impulse from the processed signal:

```text
impulse = trapz(y_processed[cycle_start:cycle_end], t[cycle_start:cycle_end])
```

Then:

```text
Mean_Impulse = mean(impulse for all detected cycles)
```

Meaning:

- average time-integrated thrust contribution per cycle
- unlike `Mean_Net`, this depends on both amplitude and time duration

### `Total_Impulse(N*s)`

Computed as:

```text
Total_Impulse = sum(impulse for all detected cycles)
```

Meaning:

- the total integrated thrust contribution across the whole analyzed recording

### `Detected_Period(s)`

Computed with FFT on the processed signal:

```text
Detected_Period = 1 / dominant_frequency
```

Meaning:

- the dominant periodic motion time scale in the file
- used internally to estimate cycle spacing and support peak detection

## Important Internal Definitions

### `peak_value`

For each detected cycle, this is the signal value at the cycle peak index.

Depends on measurement mode:

- processed mode:
  - peak value comes from the filtered/debiased signal
- original mode:
  - peak value comes from the raw signal

### `baseline_value`

For each peak, the script searches backward to find a low point before that peak and uses it as the cycle baseline.

Meaning:

- this is the local reference value used to compute net thrust for that cycle

### `net_thrust`

Defined as:

```text
net_thrust = peak_value - baseline_value
```

Meaning:

- the cycle amplitude relative to its baseline
- only cycles with positive net thrust are kept

### `impulse`

Defined as the numerical integral of the processed signal over the cycle interval:

```text
impulse = trapz(processed_signal over cycle)
```

Meaning:

- area under the processed force-time curve for that cycle
- a combined measure of force magnitude and duration

## Output Interpretation Tips

If you want to compare datasets:

- use `Mean_Net(N)` to compare typical cycle strength
- use `Std_Net(N)` to compare repeatability
- use `Mean_Impulse(N*s)` to compare average cycle output over time
- use `Total_Impulse(N*s)` to compare whole-recording output
- use `Detected_Period(s)` to compare motion speed

If you want to judge measurement quality:

- inspect the SVG panels for baseline drift and noise
- compare `Mean_Thrust`, `Mean_Drag`, and `Mean_Net`
- watch for large `Std_Net` if the cycles are inconsistent

## Notes

- SVG files are saved in the same source-folder tree under `thrust_analysis_results`
- when multiple folders are selected, each folder receives its own summary file
- the script tries to be robust to noisy data by combining baseline estimation, low-pass filtering, FFT period detection, and fallback cycle detection
