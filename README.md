# tts-article

Reproducible secondary analysis and manuscript-generation code for the AISTDS 2026 paper **Quality Constrained Policy Selection for Flow Matching Text to Speech Inference**.

The project evaluates fixed and adaptive numbers of function evaluations for F5-TTS. The central comparison is not only against the 32-NFE baseline, but also against the strong fixed 7-NFE policy. On the current locked test set, fixed 7 preserves the measured Whisper Small English WER and outperforms both tested adaptive controllers in end-to-end latency.

## Repository contents

- `work/aistds_secondary_analysis.py` — recomputes policy outcomes, paired bootstrap intervals, decision errors, latency regret, and error concentration.
- `work/build_aistds_quality_policy_paper.py` — generates the formatted AISTDS DOCX manuscript.
- `work/audit_aistds_docx.py` — checks document structure, margins, references, equations, figures, and citation-policy constraints.
- `requirements.txt` — Python dependencies for the analysis and document build.

Generated evidence, figures, rendered pages, and manuscript files are intentionally excluded from Git. This avoids committing private research data, generated speech, and large binary artifacts.

## Expected evidence files

Place the following files in one source directory:

- `raw_timing.csv`
- `asr_metrics_small_en.csv`
- `controller_decisions.csv`
- `manifest.json`

The analysis expects 36 utterances, NFE values 5, 7, 10, and 32, and the calibration/test split recorded in the evidence files.

## Setup

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Run the secondary analysis

```powershell
python work\aistds_secondary_analysis.py `
  C:\path\to\expanded_study `
  work\aistds_analysis
```

This creates the audit tables, JSON summary, and three 300-dpi figures required by the manuscript builder.

## Build and audit the manuscript

```powershell
python work\build_aistds_quality_policy_paper.py
python work\audit_aistds_docx.py
```

Use `--analysis-dir` and `--output` to override the default repository-relative paths.

The generated document is written to:

```text
outputs/AISTDS_2026_Quality_Constrained_TTS_Policy_Selection.docx
```

Before submission, replace the visible placeholders for the academic supervisor, funding, conflict of interest, and public repository URL.

## Evidence boundary

Whisper Small English provides automatic intelligibility evidence only. The current experiment does not support claims about human-rated naturalness, speaker similarity, prosody, or multilingual generalization.

The related six-utterance pilot is cited transparently in the manuscript as a separate work under review. It is not used as independent corroboration of the expanded study.
