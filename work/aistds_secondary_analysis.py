from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


SEED = 20260927
NFES = [5, 7, 10, 32]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Secondary decision-theoretic audit of the expanded F5-TTS study"
    )
    parser.add_argument("source_dir", type=Path)
    parser.add_argument("output_dir", type=Path)
    return parser.parse_args()


def minimal_passing_nfe(group: pd.DataFrame) -> int:
    errors = dict(zip(group["nfe"].astype(int), group["word_errors"].astype(int)))
    baseline = errors[32]
    for nfe in NFES:
        if errors[nfe] <= baseline:
            return nfe
    raise RuntimeError(f"No passing NFE for {group['utterance_id'].iloc[0]}")


def main() -> None:
    args = parse_args()
    source = args.source_dir.resolve()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)

    timing = pd.read_csv(source / "raw_timing.csv")
    quality = pd.read_csv(source / "asr_metrics_small_en.csv")
    decisions = pd.read_csv(source / "controller_decisions.csv")

    if not (timing["status"] == "ok").all():
        raise RuntimeError("Timing evidence contains failed rows")
    if timing.duplicated(["utterance_id", "nfe", "repeat"]).any():
        raise RuntimeError("Timing evidence contains duplicate measurement keys")

    timing_agg = timing.groupby(["utterance_id", "split", "complexity", "nfe"], as_index=False).agg(
        median_total_seconds=("total_seconds", "median"),
        median_audio_seconds=("audio_seconds", "median"),
        median_rtf=("end_to_end_rtf", "median"),
    )

    labels = (
        quality.groupby("utterance_id", group_keys=False)
        .apply(minimal_passing_nfe, include_groups=False)
        .rename("oracle_nfe")
        .reset_index()
    )
    label_frame = (
        quality[["utterance_id", "split", "complexity"]]
        .drop_duplicates()
        .merge(labels, on="utterance_id", validate="one_to_one")
    )
    calibration = label_frame[label_frame["split"] == "calibration"].copy()
    test_labels = label_frame[label_frame["split"] == "test"].copy()

    # This policy is determined only from the calibration partition. It selects
    # the largest observed passing NFE in each human-assigned text stratum.
    conservative_map = (
        calibration.groupby("complexity")["oracle_nfe"].max().astype(int).to_dict()
    )

    test = decisions.merge(
        test_labels[["utterance_id", "oracle_nfe"]],
        on="utterance_id",
        suffixes=("_recorded", ""),
        validate="one_to_one",
    )
    if "oracle_nfe_recorded" in test.columns:
        if not (test["oracle_nfe_recorded"].astype(int) == test["oracle_nfe"].astype(int)).all():
            raise RuntimeError("Recorded and recomputed oracle labels differ")
        test = test.drop(columns=["oracle_nfe_recorded"])
    test_ids = test["utterance_id"].tolist()

    policies: dict[str, dict[str, int]] = {
        "B32": {uid: 32 for uid in test_ids},
        "F5": {uid: 5 for uid in test_ids},
        "F7": {uid: 7 for uid in test_ids},
        "H-Adaptive": dict(zip(test_ids, test["h_nfe"].astype(int))),
        "ML-Adaptive": dict(zip(test_ids, test["ml_nfe"].astype(int))),
        "C-Stratum": {
            row.utterance_id: int(conservative_map[row.complexity])
            for row in test.itertuples(index=False)
        },
        "Oracle": dict(zip(test_ids, test["oracle_nfe"].astype(int))),
    }
    overhead_ms = {
        "B32": 0.0,
        "F5": 0.0,
        "F7": 0.0,
        "H-Adaptive": 0.004600035026669502,
        "ML-Adaptive": 6.381350016454235,
        "C-Stratum": 0.0,
        "Oracle": 0.0,
    }

    timing_lookup = timing_agg.set_index(["utterance_id", "nfe"])
    quality_lookup = quality.set_index(["utterance_id", "nfe"])

    def vectors(system: str):
        times, audio, errors, words = [], [], [], []
        for uid in test_ids:
            nfe = policies[system][uid]
            trow = timing_lookup.loc[(uid, nfe)]
            qrow = quality_lookup.loc[(uid, nfe)]
            times.append(float(trow["median_total_seconds"]) + overhead_ms[system] / 1000.0)
            audio.append(float(trow["median_audio_seconds"]))
            errors.append(int(qrow["word_errors"]))
            words.append(int(qrow["reference_words"]))
        return (
            np.asarray(times, dtype=float),
            np.asarray(audio, dtype=float),
            np.asarray(errors, dtype=int),
            np.asarray(words, dtype=int),
        )

    base_time, _, base_errors, base_words = vectors("B32")
    oracle_time, _, _, _ = vectors("Oracle")
    fixed7_time, _, _, _ = vectors("F7")
    rng = np.random.default_rng(SEED)

    policy_rows = []
    for system in policies:
        system_time, system_audio, system_errors, system_words = vectors(system)
        speedup = float(base_time.sum() / system_time.sum())
        corpus_wer = float(system_errors.sum() / system_words.sum())
        base_wer = float(base_errors.sum() / base_words.sum())
        delta_pp = 100.0 * (corpus_wer - base_wer)
        boot_speed, boot_delta = [], []
        for _ in range(10000):
            idx = rng.integers(0, len(test_ids), len(test_ids))
            boot_speed.append(float(base_time[idx].sum() / system_time[idx].sum()))
            boot_delta.append(
                100.0
                * (
                    system_errors[idx].sum() / system_words[idx].sum()
                    - base_errors[idx].sum() / base_words[idx].sum()
                )
            )
        selected = np.asarray([policies[system][uid] for uid in test_ids], dtype=int)
        oracle = test.set_index("utterance_id").loc[test_ids, "oracle_nfe"].to_numpy(dtype=int)
        policy_rows.append(
            {
                "system": system,
                "audio_weighted_rtf": float(system_time.sum() / system_audio.sum()),
                "speedup_vs_B32": speedup,
                "speedup_ci_low": float(np.quantile(boot_speed, 0.025)),
                "speedup_ci_high": float(np.quantile(boot_speed, 0.975)),
                "corpus_wer": corpus_wer,
                "delta_wer_pp_vs_B32": delta_pp,
                "delta_wer_ci_low_pp": float(np.quantile(boot_delta, 0.025)),
                "delta_wer_ci_high_pp": float(np.quantile(boot_delta, 0.975)),
                "controller_overhead_ms_median": overhead_ms[system],
                "total_seconds_over_oracle": float(system_time.sum() - oracle_time.sum()),
                "total_seconds_over_F7": float(system_time.sum() - fixed7_time.sum()),
                "exact_oracle_decisions": int(np.sum(selected == oracle)),
                "overallocated_decisions": int(np.sum(selected > oracle)),
                "underallocated_decisions": int(np.sum(selected < oracle)),
                "mean_excess_nfe_vs_oracle": float(np.mean(selected - oracle)),
                "selection_counts": json.dumps(
                    {str(int(key)): int(value) for key, value in sorted(Counter(selected).items())}
                ),
            }
        )

    policy_audit = pd.DataFrame(policy_rows)
    policy_audit.to_csv(output / "policy_audit.csv", index=False, encoding="utf-8-sig")

    # Error concentration for the aggressive fixed-five setting.
    qtest = quality[quality["split"] == "test"].copy()
    q5 = qtest[qtest["nfe"] == 5][
        ["utterance_id", "complexity", "word_errors", "reference_words", "reference", "hypothesis"]
    ].rename(columns={"word_errors": "errors_nfe5", "hypothesis": "hypothesis_nfe5"})
    q32 = qtest[qtest["nfe"] == 32][
        ["utterance_id", "word_errors", "hypothesis"]
    ].rename(columns={"word_errors": "errors_nfe32", "hypothesis": "hypothesis_nfe32"})
    risk = q5.merge(q32, on="utterance_id", validate="one_to_one")
    risk["extra_errors_nfe5"] = risk["errors_nfe5"] - risk["errors_nfe32"]
    risk["passes_utterance_gate"] = risk["extra_errors_nfe5"] <= 0
    risk = risk.merge(
        test[
            [
                "utterance_id",
                "word_count",
                "char_count",
                "mean_word_length",
                "long_word_count",
                "punctuation_count",
                "clause_count",
                "digit_count",
                "uppercase_token_count",
                "hyphen_count",
            ]
        ],
        on="utterance_id",
        validate="one_to_one",
    )
    risk.to_csv(output / "utterance_risk_audit.csv", index=False, encoding="utf-8-sig")

    risk_strata = risk.groupby("complexity", as_index=False).agg(
        utterances=("utterance_id", "count"),
        passing_at_nfe5=("passes_utterance_gate", "sum"),
        extra_word_errors_at_nfe5=("extra_errors_nfe5", "sum"),
        reference_words=("reference_words", "sum"),
    )
    risk_strata["pass_rate_at_nfe5"] = risk_strata["passing_at_nfe5"] / risk_strata["utterances"]
    risk_strata.to_csv(output / "risk_strata.csv", index=False, encoding="utf-8-sig")

    # Decision-level audit table for the three non-oracle adaptive policies.
    audit_rows = []
    oracle_map = policies["Oracle"]
    for system in ["H-Adaptive", "ML-Adaptive", "C-Stratum"]:
        for uid in test_ids:
            selected = policies[system][uid]
            target = oracle_map[uid]
            audit_rows.append(
                {
                    "system": system,
                    "utterance_id": uid,
                    "complexity": str(test.set_index("utterance_id").loc[uid, "complexity"]),
                    "selected_nfe": selected,
                    "oracle_nfe": target,
                    "decision": "exact" if selected == target else ("over" if selected > target else "under"),
                    "excess_nfe": selected - target,
                }
            )
    decision_audit = pd.DataFrame(audit_rows)
    decision_audit.to_csv(output / "decision_audit.csv", index=False, encoding="utf-8-sig")

    plt.rcParams.update(
        {
            "font.family": "DejaVu Sans",
            "font.size": 9,
            "text.color": "#111111",
            "axes.labelcolor": "#111111",
            "axes.edgecolor": "#555555",
            "xtick.color": "#111111",
            "ytick.color": "#111111",
        }
    )

    # Figure 1: policy frontier.
    fig, ax = plt.subplots(figsize=(7.1, 4.1), dpi=300)
    marker_map = {
        "B32": "o",
        "F5": "v",
        "F7": "s",
        "H-Adaptive": "^",
        "ML-Adaptive": "D",
        "C-Stratum": "P",
        "Oracle": "*",
    }
    offsets = {
        "B32": (5, 5),
        "F5": (5, 5),
        "F7": (10, 16),
        "H-Adaptive": (-92, -18),
        "ML-Adaptive": (-36, 14),
        "C-Stratum": (5, 5),
        "Oracle": (8, -24),
    }
    for row in policy_audit.itertuples(index=False):
        ax.scatter(
            row.speedup_vs_B32,
            row.delta_wer_pp_vs_B32,
            marker=marker_map[row.system],
            s=72 if row.system != "Oracle" else 105,
            facecolor="white" if row.system not in {"F7", "Oracle"} else "#444444",
            edgecolor="#222222",
            linewidth=1.2,
            zorder=3,
        )
        ax.annotate(
            row.system,
            (row.speedup_vs_B32, row.delta_wer_pp_vs_B32),
            xytext=offsets[row.system],
            textcoords="offset points",
        )
    ax.axhline(1.0, color="#777777", linestyle="--", linewidth=1.0, label="Provisional +1 pp WER margin")
    ax.axhline(0.0, color="#AAAAAA", linewidth=0.8)
    ax.set_xlabel("End-to-end speedup relative to 32 NFE")
    ax.set_ylabel("Corpus WER difference from 32 NFE, percentage points")
    ax.set_title("Quality-constrained policy comparison on the 18-utterance test set")
    ax.grid(True, color="#D7D7D7", alpha=0.8)
    ax.legend(frameon=False, loc="upper left")
    fig.tight_layout()
    fig.savefig(output / "fig_policy_frontier_300dpi.png", dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    # Figure 2: decision audit.
    counts = (
        decision_audit.groupby(["system", "decision"]).size().unstack(fill_value=0)
        .reindex(["H-Adaptive", "ML-Adaptive", "C-Stratum"])
    )
    for col in ["exact", "over", "under"]:
        if col not in counts.columns:
            counts[col] = 0
    fig, ax = plt.subplots(figsize=(6.8, 3.8), dpi=300)
    left = np.zeros(len(counts))
    shades = {"exact": "#333333", "over": "#999999", "under": "#FFFFFF"}
    hatches = {"exact": "", "over": "//", "under": "xx"}
    for decision in ["exact", "over", "under"]:
        values = counts[decision].to_numpy()
        ax.barh(
            counts.index,
            values,
            left=left,
            color=shades[decision],
            edgecolor="#222222",
            hatch=hatches[decision],
            label=decision.capitalize(),
        )
        left += values
    ax.set_xlim(0, 18)
    ax.set_xlabel("Test utterances")
    ax.set_title("Controller decisions relative to the per-utterance oracle")
    ax.legend(frameon=False, ncol=3, loc="lower center", bbox_to_anchor=(0.5, -0.32))
    ax.grid(True, axis="x", color="#D7D7D7", alpha=0.8)
    fig.tight_layout()
    fig.savefig(output / "fig_decision_audit_300dpi.png", dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    # Figure 3: localization of the five-NFE errors.
    order = {"simple": 0, "medium": 1, "complex": 2}
    risk_plot = risk.assign(_order=risk["complexity"].map(order)).sort_values(["_order", "utterance_id"])
    fig, ax = plt.subplots(figsize=(7.2, 3.8), dpi=300)
    colors = risk_plot["complexity"].map({"simple": "#D9D9D9", "medium": "#888888", "complex": "#333333"})
    ax.bar(risk_plot["utterance_id"], risk_plot["extra_errors_nfe5"], color=colors, edgecolor="#222222")
    ax.axhline(0, color="#555555", linewidth=0.8)
    ax.set_ylabel("Additional ASR word errors at 5 NFE vs 32 NFE")
    ax.set_xlabel("Test utterance")
    ax.set_title("Five-NFE intelligibility errors were concentrated in three complex sentences")
    ax.tick_params(axis="x", rotation=55)
    ax.grid(True, axis="y", color="#D7D7D7", alpha=0.8)
    fig.tight_layout()
    fig.savefig(output / "fig_error_concentration_300dpi.png", dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    summary = {
        "analysis_type": "post-hoc secondary decision audit",
        "analysis_date": "2026-09-27",
        "source_manifest": str(source / "manifest.json"),
        "source_rows": {
            "timing": int(len(timing)),
            "asr": int(len(quality)),
            "test_utterances": int(len(test_ids)),
        },
        "calibration_derived_conservative_mapping": conservative_map,
        "f5_error_concentration": {
            "test_utterances_passing": int(risk["passes_utterance_gate"].sum()),
            "test_utterances_total": int(len(risk)),
            "extra_word_errors": int(risk["extra_errors_nfe5"].sum()),
            "utterances_with_extra_errors": risk.loc[
                risk["extra_errors_nfe5"] > 0, "utterance_id"
            ].tolist(),
        },
        "policy_audit": policy_audit.to_dict(orient="records"),
        "claim_boundary": (
            "Whisper Small EN provides automatic intelligibility evidence only. "
            "No speaker-similarity or human-listening claim is supported."
        ),
    }
    (output / "secondary_analysis_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
