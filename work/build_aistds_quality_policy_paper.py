from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
ANALYSIS = ROOT / "work" / "aistds_analysis"
OUT = ROOT / "outputs" / "AISTDS_2026_Quality_Constrained_TTS_Policy_Selection.docx"

FIG_FRONTIER = ANALYSIS / "fig_policy_frontier_300dpi.png"
FIG_DECISIONS = ANALYSIS / "fig_decision_audit_300dpi.png"
FIG_ERRORS = ANALYSIS / "fig_error_concentration_300dpi.png"

FONT = "Times New Roman"
BLACK = RGBColor(0, 0, 0)
GRAY = "E7E6E6"
LIGHT = "F2F2F2"


def set_font(run, size=11, bold=False, italic=False):
    run.font.name = FONT
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = BLACK
    return run


def configure(doc):
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.top_margin = Mm(25.4)
    sec.bottom_margin = Mm(20)
    sec.left_margin = Mm(18)
    sec.right_margin = Mm(18)
    sec.header_distance = Mm(17)
    sec.footer_distance = Mm(17)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal.font.size = Pt(11)
    normal.font.color.rgb = BLACK
    normal.paragraph_format.line_spacing = 1.0
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.first_line_indent = Mm(5)

    for name, size, before, after in [
        ("Title", 14, 0, 8),
        ("Heading 1", 12, 10, 5),
        ("Heading 2", 11, 7, 3),
    ]:
        style = doc.styles[name]
        style.font.name = FONT
        style._element.rPr.rFonts.set(qn("w:ascii"), FONT)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = BLACK
        style.paragraph_format.first_line_indent = Mm(0)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.line_spacing = 1.0
        style.paragraph_format.keep_with_next = True

    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.first_line_indent = Mm(0)
    run = footer.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])
    set_font(run, 9)


def heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Mm(0)
    set_font(p.add_run(text), 12 if level == 1 else 11, bold=True)
    return p


def body(doc, text, *, first=True, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.first_line_indent = Mm(5 if first else 0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run(text), 11, italic=italic)
    return p


def add_labeled_paragraph(doc, label, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Mm(0)
    p.paragraph_format.line_spacing = 1.0
    set_font(p.add_run(label + " "), 11, bold=True)
    set_font(p.add_run(text), 11)
    return p


def page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Mm(0)
    p.add_run().add_break(WD_BREAK.PAGE)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_cell_width(cell, twips):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(twips))
    tc_w.set(qn("w:type"), "dxa")


def add_table(doc, caption_text, headers, rows, widths, font_size=9.5):
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.first_line_indent = Mm(0)
    cap.paragraph_format.keep_with_next = True
    cap.paragraph_format.space_before = Pt(5)
    cap.paragraph_format.space_after = Pt(3)
    set_font(cap.add_run(caption_text), 10, bold=True)

    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = False
    set_repeat_header(table.rows[0])
    for idx, value in enumerate(headers):
        cell = table.rows[0].cells[idx]
        set_cell_width(cell, widths[idx])
        set_cell_shading(cell, GRAY)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Mm(0)
        p.paragraph_format.space_after = Pt(0)
        set_font(p.add_run(value), font_size, bold=True)
    for row_index, row in enumerate(rows):
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            cell = cells[idx]
            set_cell_width(cell, widths[idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if row_index % 2:
                set_cell_shading(cell, LIGHT)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Mm(0)
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(0)
            set_font(p.add_run(str(value)), font_size)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.first_line_indent = Mm(0)
    spacer.paragraph_format.space_after = Pt(2)
    return table


def add_figure(doc, path, caption_text, width_mm=156):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Mm(0)
    p.paragraph_format.keep_with_next = True
    run = p.add_run()
    run.add_picture(str(path), width=Mm(width_mm))
    pic = run._r.xpath(".//wp:inline")
    if pic:
        doc_pr = pic[0].find(qn("wp:docPr"))
        if doc_pr is not None:
            doc_pr.set("title", path.stem)
            doc_pr.set("descr", caption_text)
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.first_line_indent = Mm(0)
    cap.paragraph_format.space_after = Pt(5)
    set_font(cap.add_run(caption_text), 10, italic=True)


def clear_paragraph(paragraph):
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def math_run(text, style="i"):
    r = OxmlElement("m:r")
    r_pr = OxmlElement("m:rPr")
    sty = OxmlElement("m:sty")
    sty.set(qn("m:val"), style)
    r_pr.append(sty)
    r.append(r_pr)
    t = OxmlElement("m:t")
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r


def equation(doc, parts, number):
    p = doc.add_paragraph()
    clear_paragraph(p)
    p_pr = p._p.get_or_add_pPr()
    p.paragraph_format.first_line_indent = Mm(0)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    tabs = OxmlElement("w:tabs")
    center = OxmlElement("w:tab")
    center.set(qn("w:val"), "center")
    center.set(qn("w:pos"), "4930")
    right = OxmlElement("w:tab")
    right.set(qn("w:val"), "right")
    right.set(qn("w:pos"), "9860")
    tabs.extend([center, right])
    p_pr.append(tabs)

    tab1 = OxmlElement("w:r")
    tab1.append(OxmlElement("w:tab"))
    p._p.append(tab1)
    omath = OxmlElement("m:oMath")
    for txt, style in parts:
        omath.append(math_run(txt, style))
    p._p.append(omath)
    tab2 = OxmlElement("w:r")
    tab2.append(OxmlElement("w:tab"))
    p._p.append(tab2)
    nr = OxmlElement("w:r")
    nr_pr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"), FONT)
    rf.set(qn("w:hAnsi"), FONT)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "22")
    nr_pr.extend([rf, sz])
    nr.append(nr_pr)
    nt = OxmlElement("w:t")
    nt.text = f"({number})"
    nr.append(nt)
    p._p.append(nr)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Mm(0)
    p.paragraph_format.left_indent = Mm(6)
    p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run(text), 11)


def add_reference(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Mm(-5)
    p.paragraph_format.left_indent = Mm(5)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(3)
    set_font(p.add_run(text), 10)


def add_front_matter(doc):
    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Mm(0)
    set_font(p.add_run("Quality Constrained Policy Selection for Flow Matching Text to Speech Inference"), 14, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Mm(0)
    set_font(p.add_run("Anton Popov¹ and [Academic Supervisor Name]²"), 11, bold=True)
    body(doc, "¹State University of Information and Communication Technologies, Institute of Information Technologies, Kyiv, Ukraine", first=False, align=WD_ALIGN_PARAGRAPH.CENTER)
    body(doc, "²[Academic Supervisor Affiliation]", first=False, align=WD_ALIGN_PARAGRAPH.CENTER)
    body(doc, "Corresponding author: popov.a.o.ua@gmail.com · ORCID 0009-0006-7557-094X", first=False, align=WD_ALIGN_PARAGRAPH.CENTER)

    heading(doc, "Abstract", 1)
    body(doc, (
        "Iterative flow matching makes the number of function evaluations an explicit latency control for neural text-to-speech systems. "
        "The central deployment question is not whether fewer evaluations run faster, but which policy minimizes latency without exceeding an intelligibility loss budget. "
        "This paper presents a quality constrained policy audit for the unchanged public F5-TTS v1 Base system. Thirty six original English utterances were stratified as simple, medium, or complex and split before analysis into 18 calibration and 18 locked test items. Each item was synthesized at 5, 7, 10, and 32 function evaluations using two randomized timed repeats after two warmups, producing 288 end to end measurements. One saved output per condition was transcribed by Whisper Small English, producing 144 automatic intelligibility observations. On the locked test, fixed 7 evaluations achieved an audio weighted real time factor of 0.0473 and a 2.70 times speedup over the 32 evaluation baseline while retaining the same 3.03 percent corpus word error rate. Fixed 5 evaluations was faster at 3.20 times but increased word error rate to 5.19 percent and failed the provisional non inferiority gate. Two adaptive controllers preserved baseline word error rate but were slower than fixed 7: a deterministic heuristic achieved 2.62 times speedup and a random forest achieved 2.42 times after 6.38 ms median controller overhead. A new secondary decision audit showed that both adaptive systems overallocated 12 of 18 locked test utterances. All five excess recognized word errors at 5 evaluations were concentrated in three complex items. An oracle retained baseline word error rate at 3.09 times speedup, demonstrating headroom while also showing that present text features do not capture it. The defensible deployment choice for this evidence set is therefore fixed 7, not the tested adaptive controllers. The paper contributes a reproducible policy selection contract, decision regret analysis, and a transparent boundary between automatic intelligibility evidence and unmeasured perceptual quality."
    ))
    add_labeled_paragraph(doc, "Keywords", "neural text to speech, flow matching, adaptive inference, policy selection, real time factor, word error rate, reproducible evaluation")


def parse_args():
    parser = argparse.ArgumentParser(description="Build the AISTDS 2026 DOCX manuscript")
    parser.add_argument("--analysis-dir", type=Path, default=ANALYSIS)
    parser.add_argument("--output", type=Path, default=OUT)
    return parser.parse_args()


def build(analysis_dir: Path = ANALYSIS, output: Path = OUT):
    analysis_dir = analysis_dir.resolve()
    output = output.resolve()
    fig_frontier = analysis_dir / "fig_policy_frontier_300dpi.png"
    fig_decisions = analysis_dir / "fig_decision_audit_300dpi.png"
    fig_errors = analysis_dir / "fig_error_concentration_300dpi.png"
    for path in [fig_frontier, fig_decisions, fig_errors]:
        if not path.exists():
            raise FileNotFoundError(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    configure(doc)
    add_front_matter(doc)

    page_break(doc)
    heading(doc, "Introduction", 1)
    body(doc, (
        "Modern neural text-to-speech systems increasingly obtain natural sounding speech through iterative generative procedures. Diffusion and continuous normalizing flow formulations replace a direct acoustic mapping with a trajectory that transforms a simple source distribution into speech. Flow matching provides a particularly convenient training objective for such trajectories and has become an important basis for expressive and zero shot synthesis (Lipman et al., 2023; Popov et al., 2021). F5-TTS applies this idea in a fully non autoregressive architecture and exposes the number of function evaluations, abbreviated NFE, as an inference parameter (Chen et al., 2025). This parameter directly affects how many times the neural vector field is evaluated and therefore creates an actionable latency control."
    ))
    body(doc, (
        "Reducing NFE is attractive for interactive information technologies that process voice information. A conversational agent, accessibility reader, navigation interface, or embedded voice service must respond within a bounded time while preserving intelligibility. Yet a low step count cannot be selected from latency alone. A system that speaks quickly but changes numbers, omits words, or corrupts technical expressions has not improved the application. Conversely, a conservative controller can preserve quality by assigning extra evaluations to almost every input, but then the controller adds complexity without delivering a useful speed advantage. The engineering objective is thus constrained policy selection rather than unconstrained acceleration."
    ))
    body(doc, (
        "Previous TTS research offers fast parallel architectures such as FastSpeech 2 and VITS, diffusion based acoustic generation, and increasingly efficient flow models (Kim et al., 2021; Ren et al., 2021; Tan et al., 2024). Related work on numerical solvers and adaptive time stepping shows that computational effort need not be uniform across samples (Chen et al., 2024; Lu et al., 2022). However, a sample specific controller is useful only if it outperforms a strong fixed fast setting under the same quality criterion and timing boundary. Comparing an adaptive policy only with a 32 step default can hide this requirement."
    ))
    body(doc, (
        "This paper asks three questions. First, which tested fixed NFE setting satisfies a provisional automatic intelligibility constraint on a locked test set? Second, do deterministic or learned text side controllers improve upon that fixed setting after their overhead is counted? Third, where are low NFE failures concentrated, and what does that concentration imply for the design of an information technology for neural voice processing? The answers are deliberately limited to automatic intelligibility measured by a frozen speech recognizer. No human naturalness, speaker similarity, prosody preservation, or multilingual conclusion is inferred."
    ))
    heading(doc, "Contributions", 2)
    bullet(doc, "A controlled evidence set with 36 original utterances, 288 end to end timing rows, 144 frozen automatic speech recognition evaluations, and a calibration locked test split.")
    bullet(doc, "A quality constrained comparison of three fixed policies, two implementable adaptive policies, a calibration derived conservative policy, and an oracle lower bound.")
    bullet(doc, "A new decision audit that separates exact allocation, harmless over allocation, unsafe under allocation, controller overhead, and latency regret relative to the oracle and fixed 7.")
    bullet(doc, "A transparent negative result: fixed 7 is the best supported measured deployment point even though oracle allocation confirms that sample specific headroom exists.")

    page_break(doc)
    heading(doc, "Relation to the Under Review Manuscript", 1)
    body(doc, (
        "An earlier manuscript by the first author, currently under review, reported a six utterance pilot of fixed step F5-TTS acceleration and a post hoc feasibility mapping (Popov, 2026). It established that low NFE inference was computationally promising but did not demonstrate a successful deployed adaptive controller. The present paper cites that manuscript explicitly because the research program, base synthesizer, and broad motivation are related. The under review manuscript is not treated as independent confirmation of the present findings."
    ))
    body(doc, (
        "The current study has a distinct analytical object. Instead of asking whether adaptive allocation is plausible, it asks whether any tested adaptive policy is rational to deploy when fixed 7 is included as a competitive comparator. The data are newly expanded from six to 36 utterances, the protocol produces 288 timed generations rather than 72, the automatic evaluator changes from a tiny English recognizer to Whisper Small English, and the test is locked before policy assessment. The secondary analysis introduced here was not the focus of the previous manuscript: it measures policy decisions against utterance level oracle labels, quantifies excess NFE, and localizes the word errors that cause fixed 5 to fail."
    ))
    add_table(doc, "Table 1 Distinction between the related manuscripts", ["Dimension", "Earlier under review manuscript", "Present AISTDS manuscript"], [
        ["Primary question", "Is adaptive low step inference feasible?", "Which policy should be selected under a quality constraint?"],
        ["Speech material", "6 English utterances", "36 new original English utterances"],
        ["Timed generations", "72", "288"],
        ["Automatic evaluator", "Faster Whisper Tiny English", "Whisper Small English"],
        ["Adaptive evidence", "Post hoc feasibility mapping", "Executed heuristic and random forest on locked test"],
        ["New analysis", "Latency and pilot WER tradeoff", "Decision error, excess NFE, regret, and risk concentration"],
        ["Principal conclusion", "Headroom justifies a larger study", "Fixed 7 dominates tested adaptive policies under measured WER"],
    ], [1900, 3860, 4100], 9.2)
    body(doc, (
        "This separation matters for research integrity. The new paper neither copies the earlier pilot as a new experiment nor disguises the related submission. It uses the earlier work to motivate the need for a stronger comparison, then reports a larger evidence package and a new conclusion that is less favorable to adaptivity. Such a result is scientifically useful because it prevents an information technology architecture from adopting a controller whose decision cost and conservative assignments do not improve the observed frontier."
    ))

    page_break(doc)
    heading(doc, "Related Work", 1)
    heading(doc, "Neural and Flow Based Speech Synthesis", 2)
    body(doc, (
        "Non autoregressive speech synthesis reduced the sequential bottleneck of earlier acoustic models. FastSpeech 2 incorporated duration, pitch, and energy prediction, while VITS combined variational inference, adversarial training, and normalizing flows in an end to end waveform model (Kim et al., 2021; Ren et al., 2021). Grad TTS formulated acoustic generation as a gradual transformation governed by a stochastic differential equation and exposed a quality computation tradeoff (Popov et al., 2021). NaturalSpeech extended end to end TTS toward human level naturalness on benchmark conditions, illustrating the increasing importance of comprehensive evaluation rather than a single acoustic proxy (Tan et al., 2024)."
    ))
    body(doc, (
        "Flow matching directly regresses a time dependent vector field along a chosen probability path (Lipman et al., 2023). F5-TTS uses a diffusion transformer and a flow matching objective for fully non autoregressive zero shot synthesis, with an inference procedure in which the requested NFE controls repeated model evaluations (Chen et al., 2025). This open implementation makes it possible to study policy selection without retraining the base acoustic model. The present work treats the synthesizer as fixed and changes only the computation assigned at inference."
    ))
    heading(doc, "Acceleration and Adaptive Computation", 2)
    body(doc, (
        "Acceleration can be obtained by altering model architecture, training objectives, distillation, numerical integration, or inference schedules. DPM Solver demonstrates that specialized high order solvers can reduce diffusion model evaluations without retraining the score model (Lu et al., 2022). Adaptive time stepping similarly allocates solver effort according to the local trajectory rather than imposing a uniform grid (Chen et al., 2024). These methods operate inside the generative trajectory. A text side NFE controller instead acts before synthesis and must infer risk from the input. Its appeal is implementation simplicity, but its prediction problem can be harder because linguistic features are only indirect indicators of acoustic failure."
    ))
    heading(doc, "Evaluation of Speech Generation", 2)
    body(doc, (
        "Speech quality is multidimensional. Word error rate can reveal substitutions, deletions, and insertions, but it is not a measure of naturalness or speaker identity. The VoiceMOS and UTMOS studies show both the value and limitations of automatic predictors for perceptual assessment (Huang et al., 2022; Saeki et al., 2022). Reviews of speech analysis and multimodal audio research likewise emphasize that metric choice must follow the intended construct (Mehrish et al., 2023; O’Shaughnessy, 2023). For this reason, the present claim is restricted to automatic intelligibility evidence from a fixed recognizer."
    ))
    body(doc, (
        "The policy selection problem also requires paired comparisons. Each utterance is synthesized under every candidate NFE, so within utterance ratios and differences are more informative than unrelated sample averages. Bootstrap intervals preserve this paired structure. A policy is not declared acceptable merely because its point estimate is close to baseline; its upper uncertainty bound must remain within a declared margin. This approach follows the broader principle that evaluation gates should be specified before deployment decisions and should not be replaced by impressions formed after observing the results."
    ))

    page_break(doc)
    heading(doc, "Quality Constrained Policy Selection", 1)
    heading(doc, "System Boundary", 2)
    body(doc, (
        "Let xᵢ be normalized text, r the fixed reference utterance, k the NFE budget, and Gθ the unchanged F5-TTS synthesizer. The measured time Tᵢ includes acoustic generation, vocoding, device synchronization, and WAV writing. For adaptive systems it also includes the time needed to compute text features and choose k. Model loading and warmup are excluded because they are one time service initialization costs rather than per request inference costs. Audio duration Dᵢ is measured from the saved output."
    ))
    equation(doc, [("RTF", "p"), ("ᵢ(k) = T", "i"), ("ᵢ(k) / D", "i"), ("ᵢ(k)", "i")], 1)
    body(doc, "An audio weighted corpus RTF is computed as the sum of measured synthesis time divided by the sum of generated audio duration. The speedup of a policy π relative to the 32 evaluation baseline is therefore", first=False)
    equation(doc, [("S(π) = RTF(B32) / RTF(π)", "p")], 2)
    heading(doc, "Automatic Intelligibility Constraint", 2)
    body(doc, (
        "The frozen evaluator transcribes one saved output per utterance and NFE. Reference and hypothesis strings use the same lowercase and punctuation normalization. Corpus WER is the ratio of substitutions, deletions, and insertions to reference words. For each utterance, the oracle admissible budget is the smallest tested k whose word error count does not exceed the error count of the same utterance at 32 NFE."
    ))
    equation(doc, [("k", "i"), ("*", "p"), (" = min {k ∈ {5, 7, 10, 32} : e", "p"), ("ᵢ(k) ≤ e", "i"), ("ᵢ(32)}", "i")], 3)
    body(doc, (
        "The provisional deployment gate compares a candidate with B32. The upper limit of the paired bootstrap 95 percent interval for the WER difference must not exceed plus one percentage point, and the lower limit of the speedup interval must exceed one. The one point margin is an engineering threshold for this study, not a universal perceptual equivalence margin."
    ))
    equation(doc, [("Accept(π) ⇔ U", "p"), ("95", "p"), ("[ΔWER(π)] ≤ 1 pp ∧ L", "p"), ("95", "p"), ("[S(π)] > 1", "p")], 4)
    heading(doc, "Policies", 2)
    add_table(doc, "Table 2 Evaluated inference policies", ["Policy", "Rule", "Purpose"], [
        ["B32", "Always use 32 NFE", "Quality reference"],
        ["F5", "Always use 5 NFE", "Aggressive fixed baseline"],
        ["F7", "Always use 7 NFE", "Strong fixed fast baseline"],
        ["H Adaptive", "Simple 5, medium 7, complex 10", "Transparent heuristic"],
        ["ML Adaptive", "Random forest over ten text features with fallback to 10", "Lightweight learned controller"],
        ["C Stratum", "Calibration maximum passing NFE within each stratum", "Conservative post hoc audit"],
        ["Oracle", "Smallest locked test NFE matching B32 error count", "Attainable lower bound, not deployable"],
    ], [1500, 5200, 3160], 9.2)
    body(doc, (
        "The heuristic maps the three declared complexity strata directly. The learned controller uses word count, character count, mean word length, number of long words, clause count, punctuation count, digit count, uppercase token count, hyphen count, and question count. It is a random forest with 300 trees, maximum depth three, minimum leaf size two, and balanced class weights. A low confidence prediction falls back to 10 NFE. C Stratum is included only as a secondary conservative comparator because it uses the maximum observed passing NFE in each calibration stratum."
    ))

    page_break(doc)
    heading(doc, "Experimental Design", 1)
    heading(doc, "Speech Material and Split", 2)
    body(doc, (
        "The evidence set contains 36 original English utterances written for the experiment. Twelve items were assigned to each of three predefined strata. Simple items use short common words and uncomplicated syntax. Medium items add longer phrases, punctuation, or a moderate clause structure. Complex items include combinations of decimals, abbreviations, uppercase tokens, hyphenated forms, rare technical terms, or multiple clauses. A fixed random seed split the set into 18 calibration and 18 locked test utterances, preserving six items from each stratum in both partitions."
    ))
    body(doc, (
        "The calibration partition was used to derive utterance level admissible labels, fit the random forest, and freeze the fallback. The locked test partition was not used to alter features, class rules, hyperparameters, or thresholds. The audit nevertheless remains exploratory because the number of calibration items is small and the utterances are constructed rather than sampled from a natural traffic distribution. The term locked describes procedural separation, not a claim of population representativeness."
    ))
    heading(doc, "Implementation and Timing", 2)
    add_table(doc, "Table 3 Reproducibility manifest", ["Component", "Pinned setting"], [
        ["Repository", "SWivid F5-TTS commit 9c614e9657089213efc6a7421b30630be138a3f5"],
        ["Package and model", "f5 tts 1.1.22 and F5TTS v1 Base"],
        ["Hardware", "NVIDIA GeForce RTX 5080 with 16 303 MiB memory"],
        ["Software", "Windows, Python 3.12, PyTorch 2.9.1 with CUDA 13.0"],
        ["Synthesis", "Euler sampling, sway minus 1, CFG 2.0, 24 kHz, fixed seed"],
        ["NFE values", "5, 7, 10, and 32"],
        ["Timing", "Two warmups and two randomized measured repeats per condition"],
        ["Automatic evaluator", "Whisper Small English snapshot e8727524f962ee844a7319d92be39ac1bd25655a"],
    ], [2600, 7260], 9.2)
    body(doc, (
        "Each utterance was synthesized under all four NFE conditions. Two warmup calls were completed before measured inference, followed by two timed repeats in randomized order. CUDA synchronization was performed at timing boundaries. Each measured repeat was written to WAV so that storage overhead was consistently included. No generation failed, yielding 36 multiplied by four multiplied by two, or 288, complete timing rows. One designated output per utterance and NFE was retained for recognition, yielding 144 ASR rows."
    ))
    body(doc, (
        "The official English reference prompt and the same reference audio were used throughout. Audio was stored at 24 kHz. A TorchCodec dynamic library incompatibility in the environment was bypassed with a soundfile based WAV loader; the acoustic model, vocoder, sampling path, and evaluation data were otherwise unchanged. This workaround is recorded because reproducibility includes infrastructure failures that affect whether the experiment can be executed."
    ))
    heading(doc, "Statistical Analysis", 2)
    body(doc, (
        "Latency is summarized by audio weighted RTF. Speedup uncertainty is obtained with paired bootstrap resampling over the 18 locked test utterances. WER difference intervals are also paired by utterance, preserving the fact that every policy is evaluated on the same text. The secondary audit counts decisions equal to, greater than, or less than the oracle admissible NFE. Greater decisions are classified as over allocation and lower decisions as under allocation. Because NFE values are discrete, excess allocation is reported both as counts and as mean NFE difference."
    ))
    equation(doc, [("R", "i"), ("T", "p"), ("(π) = T", "p"), ("ᵢ(π(x", "i"), ("ᵢ)) − T", "i"), ("ᵢ(k", "i"), ("ᵢ* )", "i")], 5)
    body(doc, "The sum of RᵢT is latency regret relative to the oracle. A second regret measure replaces the oracle with F7, which is the strongest fixed policy that passes the observed WER gate.", first=False)

    page_break(doc)
    heading(doc, "Results", 1)
    heading(doc, "Latency and Automatic Intelligibility", 2)
    body(doc, (
        "On the locked test, B32 achieved an audio weighted RTF of 0.1277 and a corpus WER of 3.03 percent. F5 reduced RTF to 0.0399, corresponding to a 3.20 times speedup with a paired bootstrap interval from 3.03 to 3.37. Its WER rose to 5.19 percent, a difference of 2.16 percentage points. Although the paired WER interval included zero because only three utterances were affected, its upper limit was 4.82 percentage points, so F5 failed the provisional plus one point gate."
    ))
    body(doc, (
        "F7 achieved RTF 0.0473 and a 2.70 times speedup with a 95 percent interval from 2.58 to 2.82. Its corpus WER was identical to B32 at 3.03 percent. H Adaptive also retained 3.03 percent WER and reached 2.62 times speedup, while ML Adaptive retained 3.03 percent WER and reached 2.42 times speedup. The random forest controller itself required 6.38 ms median decision time. The heuristic overhead was negligible at 0.0046 ms, but its NFE assignments were more conservative than necessary."
    ))
    add_table(doc, "Table 4 Locked test policy outcomes", ["Policy", "RTF", "Speedup and 95 percent interval", "WER", "WER difference", "Gate"], [
        ["B32", "0.1277", "1.00", "3.03%", "0.00 pp", "Reference"],
        ["F5", "0.0399", "3.20 [3.03, 3.37]", "5.19%", "+2.16 pp", "Fail"],
        ["F7", "0.0473", "2.70 [2.58, 2.82]", "3.03%", "0.00 pp", "Pass"],
        ["H Adaptive", "0.0488", "2.62 [2.51, 2.76]", "3.03%", "0.00 pp", "Pass"],
        ["ML Adaptive", "0.0527", "2.42 [2.31, 2.55]", "3.03%", "0.00 pp", "Pass"],
        ["C Stratum", "0.0779", "1.64 [1.32, 2.19]", "3.03%", "0.00 pp", "Pass"],
        ["Oracle", "0.0414", "3.09 [2.95, 3.23]", "3.03%", "0.00 pp", "Diagnostic"],
    ], [1250, 1000, 2500, 1100, 1500, 1260], 8.7)
    add_figure(doc, fig_frontier, "Figure 1 Quality constrained speed frontier on the locked test set. F7 is the fastest measured non oracle policy at baseline corpus WER; error bars show paired bootstrap speedup intervals.")
    heading(doc, "Policy Selection Outcome", 2)
    body(doc, (
        "The result changes the practical interpretation of adaptivity. If B32 were the only comparator, both adaptive controllers would appear successful because they preserve measured WER while providing more than 2.4 times speedup. Once F7 is included, neither adaptive policy improves the frontier. H Adaptive is 0.193 s slower than F7 across the locked test set, and ML Adaptive is 0.730 s slower. The latter difference reflects both controller overhead and assignment of 10 NFE to twelve utterances. C Stratum is even more conservative, adding 4.145 s relative to F7."
    ))
    body(doc, (
        "The oracle reaches a 3.09 times speedup at the same corpus WER as B32, so the negative adaptive result is not evidence that sample specific allocation is impossible. It shows that the available calibration data and text features do not locate the headroom accurately enough. A future adaptive system must therefore be evaluated against F7, not only against B32, and it must approach the oracle without increasing unsafe under allocation."
    ))

    page_break(doc)
    heading(doc, "Decision Audit and Error Concentration", 1)
    heading(doc, "Allocation Errors", 2)
    body(doc, (
        "The oracle selected 5 NFE for 15 locked test utterances and 7 NFE for three. No item required 10 or 32 NFE to match its own B32 recognized word error count. F5 exactly matched the oracle for 15 items but underallocated the three failures. F7 exactly matched three items and overallocated the remaining 15. This is an important asymmetry: F5 uses less computation but creates automatic intelligibility errors, whereas F7 spends two extra evaluations on many easy inputs yet avoids the observed penalty."
    ))
    body(doc, (
        "H Adaptive, ML Adaptive, and C Stratum each made six exact and twelve overallocated decisions with no underallocation. Their mean excess NFE relative to the oracle was 2.0, 3.0, and 9.33 respectively. Thus, none of the tested adaptive systems failed because it was too aggressive. They failed to beat F7 because they were too cautious for the actual locked test distribution. The conservative calibration maximum policy illustrates the cost of using a small stratum sample to protect against a worst observed case: one difficult calibration example can raise the budget for an entire class."
    ))
    add_figure(doc, fig_decisions, "Figure 2 Locked test allocation decisions relative to the utterance level oracle. Under allocation occurs only for F5; adaptive policies avoid under allocation but overallocate twelve of eighteen items.")
    heading(doc, "Concentration of Low Step Risk", 2)
    body(doc, (
        "All five additional recognized word errors produced by F5 were concentrated in three complex utterances labelled c07, c11, and c12. Item c07 contains an uppercase hardware name, digits, and a hyphenated expression and contributed one additional error. Item c11 contains a decimal, multiple hyphens, and several clauses and contributed three additional errors relative to B32. Item c12 contains multiple clauses and hyphenated forms and contributed one additional error. Every simple and medium locked test item passed at 5 NFE, whereas only three of six complex items did."
    ))
    add_figure(doc, fig_errors, "Figure 3 Concentration of excess recognized word errors at 5 NFE. All observed degradation is confined to three complex locked test utterances.")
    body(doc, (
        "The concentration suggests a two stage deployment architecture. A default 5 NFE path could be considered only if a high recall risk screen reliably routes the small hazardous subset to 7 NFE or higher. The present three level complexity label is not sufficient because half of the complex items still pass at 5 NFE and blanket escalation wastes computation. More informative features may include tokenizer fragmentation, grapheme phoneme ambiguity, normalized number expansion, abbreviation expansion, punctuation pattern, and uncertainty from a lightweight pronunciation or ASR proxy."
    ))
    body(doc, (
        "The same evidence also warns against a large language model controller being adopted simply because it can produce richer explanations. A remote or large controller has a stricter break even requirement than the random forest. Its total decision time must be smaller than the synthesis time saved relative to F7, and its routing errors must remain inside the quality gate. A compact classifier or deterministic preprocessor may be a better component if the risk factors are largely lexical and structural."
    ))

    page_break(doc)
    heading(doc, "Implications for Voice Information Technology", 1)
    heading(doc, "Deployment Recommendation", 2)
    body(doc, (
        "For the measured hardware, model, reference voice, and English test material, fixed 7 NFE is the appropriate default policy. It reduces audio weighted RTF by approximately 63 percent relative to B32, has a speedup interval comfortably above one, and matches B32 corpus WER. This recommendation is narrower than a claim of perceptual equivalence. The system should expose a policy version, retain a 32 NFE diagnostic mode, and log selected NFE, controller time, synthesis time, output duration, and failure state for every request."
    ))
    body(doc, (
        "An adaptive controller should remain an experimental feature until it passes three tests. First, it must be calibrated on a substantially larger and more natural sample. Second, it must outperform fixed 7 end to end, including feature extraction and decision overhead. Third, the WER gate must be complemented by speaker similarity and blinded listening. A deployment can use shadow mode: the controller predicts an NFE, but production synthesis stays at 7 while the system collects disagreement and risk data. This separates data acquisition from user facing quality risk."
    ))
    heading(doc, "Information Technology Architecture", 2)
    add_table(doc, "Table 5 Recommended processing modules", ["Module", "Input", "Output", "Audit field"], [
        ["Text normalization", "Raw text", "Canonical text and expansions", "Normalizer version"],
        ["Risk analysis", "Canonical text", "Risk score and factors", "Feature vector and latency"],
        ["Policy selection", "Risk score and service mode", "NFE budget", "Policy version and confidence"],
        ["Flow synthesis", "Text, reference, and NFE", "Waveform", "RTF, device, model hash"],
        ["Quality monitor", "Sampled waveform and text", "ASR and similarity indicators", "Metric versions"],
        ["Fallback manager", "Invalid output or high risk", "Escalated NFE or retry", "Reason and added latency"],
    ], [1800, 2250, 2700, 3110], 9.0)
    body(doc, (
        "This modular architecture aligns the research with the broader topic of information technology for neural voice processing. The contribution is not only a model comparison but a decision layer with explicit inputs, outputs, timing boundaries, evidence gates, and fallbacks. Such boundaries are essential for maintainability: a later synthesizer, recognizer, or controller can be replaced while the policy contract and audit schema remain stable."
    ))
    heading(doc, "Threats to Validity", 2)
    body(doc, (
        "Internal validity is supported by a pinned implementation, randomized repeats, synchronized GPU timing, a uniform I O boundary, complete rows, and a locked test split. Remaining internal threats include only two timed repeats, one selected recognition output per condition, and sensitivity of low latency measurements to the Windows execution environment. The paired design reduces between sentence variance but does not remove temporal system noise."
    ))
    body(doc, (
        "Construct validity is the largest limitation. Whisper Small English error counts measure one form of intelligibility. They do not establish naturalness, speaker similarity, emotional fidelity, prosody, or preference. Automatic MOS predictors should not be substituted uncritically for listeners because their behavior depends on training distributions (Huang et al., 2022; Saeki et al., 2022). A future evaluation should include at least two hundred randomized listening judgments per critical comparison, a speaker embedding similarity measure, and predefined non inferiority analyses."
    ))
    body(doc, (
        "External validity is limited to one English reference speaker, one GPU, one model checkpoint, and constructed texts. The six test items per stratum are insufficient to estimate rare failure rates. The policy may behave differently for Ukrainian, code switching, expressive prompts, long form synthesis, other voices, or hardware with a different ratio of controller to acoustic latency. The oracle is diagnostic and should not be read as an achievable online result."
    ))

    page_break(doc)
    heading(doc, "Conclusion", 1)
    body(doc, (
        "This paper reframed adaptive flow matching TTS as a quality constrained policy selection problem. The expanded experiment confirmed that lower NFE substantially reduces latency, but it also demonstrated why a 32 step baseline is not enough for judging a controller. On the locked test, fixed 7 NFE achieved a 2.70 times speedup and the same 3.03 percent corpus WER as B32. Both executed adaptive controllers preserved that WER but were slower than fixed 7. The new decision audit traced this result to conservative over allocation rather than unsafe aggressiveness."
    ))
    body(doc, (
        "The oracle result and error concentration analysis preserve a productive research direction. Fifteen of eighteen utterances matched B32 recognition at 5 NFE, and all five excess word errors were concentrated in three complex items. The next research step is therefore not to increase controller complexity indiscriminately. It is to collect more risk examples, improve features that reflect pronunciation and tokenization difficulty, and evaluate a high recall router against fixed 7 with complete overhead accounting."
    ))
    body(doc, (
        "The supported practical conclusion is intentionally conservative. Fixed 7 is the best deployment point in the present evidence. Adaptive allocation remains a hypothesis with measurable oracle headroom, not a demonstrated production advantage. This distinction, together with transparent citation of the earlier under review pilot, makes the new analysis suitable as a separate contribution rather than a repackaging of prior work."
    ))

    heading(doc, "Data and Code Availability", 2)
    body(doc, (
        "The research evidence package contains the experiment manifest, input texts, split identifiers, raw timing rows, automatic recognition metrics, policy decisions, analysis scripts, and integrity hashes. A public repository address will be inserted after the authors confirm redistribution permissions for generated audio and third party model assets: [PUBLIC REPOSITORY URL]. Generated audio may be shared separately where licensing permits; the tabular evidence and hashes allow the reported calculations to be audited without redistributing checkpoints."
    ))
    heading(doc, "Funding", 2)
    body(doc, "[Funding statement to be completed before submission.]", first=False)
    heading(doc, "Conflict of Interest", 2)
    body(doc, "[Conflict of interest statement to be completed before submission.]", first=False)
    heading(doc, "Author Contributions", 2)
    body(doc, "Anton Popov: conceptualization, software, investigation, data curation, formal analysis, visualization, and writing. [Academic Supervisor Name]: supervision, methodology review, validation, and manuscript review. Confirm roles before submission.", first=False)
    heading(doc, "Use of Generative Artificial Intelligence", 2)
    body(doc, (
        "Generative artificial intelligence tools were used to assist with language editing, document formatting, and code review. The authors are responsible for the research design, verified the numerical results against the evidence package, reviewed the references, and accept responsibility for the final manuscript. No generated text was treated as experimental evidence."
    ), first=False)

    page_break(doc)
    heading(doc, "References", 1)
    references = [
        "Ahmad, H. A., & Rashid, T. A. (2024). Planning the development of artificial intelligence based text to speech synthesis for a low resource language. Journal of King Saud University Computer and Information Sciences, 36(7), 102131. https://doi.org/10.1016/j.jksuci.2024.102131",
        "Chen, Y., He, F., Fu, S., Tian, X., & Tao, D. (2024). Adaptive time stepping for numerical integration using reinforcement learning. In Proceedings of the Fortieth Conference on Uncertainty in Artificial Intelligence (Vol. 244, pp. 685–697). PMLR. https://proceedings.mlr.press/v244/chen24h.html",
        "Chen, Y., Niu, Z., Ma, Z., Deng, K., Wang, C., Zhao, J., Yu, K., & Chen, X. (2025). F5 TTS: A fairytaler that fakes fluent and faithful speech with flow matching. In Proceedings of the 63rd Annual Meeting of the Association for Computational Linguistics (Vol. 1, pp. 6255–6271). Association for Computational Linguistics. https://doi.org/10.18653/v1/2025.acl-long.313",
        "Huang, W. C., Cooper, E., Tsao, Y., Wang, H. M., Toda, T., & Yamagishi, J. (2022). The VoiceMOS Challenge 2022. In Proceedings of Interspeech 2022 (pp. 4536–4540). https://doi.org/10.21437/Interspeech.2022-970",
        "Kim, J., Kong, J., & Son, J. (2021). Conditional variational autoencoder with adversarial learning for end to end text to speech. In Proceedings of the 38th International Conference on Machine Learning (Vol. 139, pp. 5530–5540). PMLR. https://proceedings.mlr.press/v139/kim21f.html",
        "Lipman, Y., Chen, R. T. Q., Ben Hamu, H., Nickel, M., & Le, M. (2023). Flow matching for generative modeling. In The Eleventh International Conference on Learning Representations. https://openreview.net/forum?id=PqvMRDCJT9t",
        "Liu, R., Sisman, B., Lin, Y., & Li, H. (2021). FastTalker: A neural text to speech architecture with shallow and group autoregression. Neural Networks, 141, 306–314. https://doi.org/10.1016/j.neunet.2021.04.016",
        "Lu, C., Zhou, Y., Bao, F., Chen, J., Li, C., & Zhu, J. (2022). DPM Solver: A fast ODE solver for diffusion probabilistic model sampling in around 10 steps. In Advances in Neural Information Processing Systems (Vol. 35, pp. 5775–5787). https://proceedings.neurips.cc/paper_files/paper/2022/hash/260a14acce2a89dad36adc8eefe7c59e-Abstract-Conference.html",
        "Mehrish, A., Majumder, N., Bharadwaj, R., Mihalcea, R., & Poria, S. (2023). A review of deep learning techniques for speech processing. Information Fusion, 99, 101869. https://doi.org/10.1016/j.inffus.2023.101869",
        "O’Shaughnessy, D. (2023). Review of analysis methods for speech applications. Speech Communication, 151, 64–75. https://doi.org/10.1016/j.specom.2023.05.008",
        "Popov, A. (2026). Toward adaptive AI guided step allocation for faster flow matching text to speech [Manuscript submitted for publication]. State University of Information and Communication Technologies.",
        "Popov, V., Vovk, I., Gogoryan, V., Sadekova, T., & Kudinov, M. (2021). Grad TTS: A diffusion probabilistic model for text to speech. In Proceedings of the 38th International Conference on Machine Learning (Vol. 139, pp. 8599–8608). PMLR. https://proceedings.mlr.press/v139/popov21a.html",
        "Radford, A., Kim, J. W., Xu, T., Brockman, G., McLeavey, C., & Sutskever, I. (2023). Robust speech recognition via large scale weak supervision. In Proceedings of the 40th International Conference on Machine Learning (Vol. 202, pp. 28492–28518). PMLR. https://proceedings.mlr.press/v202/radford23a.html",
        "Ren, Y., Hu, C., Tan, X., Qin, T., Zhao, S., Zhao, Z., & Liu, T. Y. (2021). FastSpeech 2: Fast and high quality end to end text to speech. In International Conference on Learning Representations. https://openreview.net/forum?id=piLPYqxtWuA",
        "Saeki, T., Xin, D., Nakata, W., Koriyama, T., Takamichi, S., & Saruwatari, H. (2022). UTMOS: UTokyo Sarulab system for VoiceMOS Challenge 2022. In Proceedings of Interspeech 2022 (pp. 4521–4525). https://doi.org/10.21437/Interspeech.2022-439",
        "Tan, X., Chen, J., Liu, H., Cong, J., Zhang, C., Liu, Y., Wang, X., Leng, Y., Yi, Y., He, L., Zhao, S., Qin, T., Soong, F., & Liu, T. Y. (2024). NaturalSpeech: End to end text to speech synthesis with human level quality. IEEE Transactions on Pattern Analysis and Machine Intelligence, 46(6), 4234–4245. https://doi.org/10.1109/TPAMI.2024.3356232",
        "Yasuda, Y., Wang, X., & Yamagishi, J. (2021). Investigation of enhanced Tacotron text to speech synthesis systems with self attention for pitch accent language. Computer Speech and Language, 67, 101183. https://doi.org/10.1016/j.csl.2020.101183",
        "Zhang, H., Wu, Z., Xing, Z., Shao, J., & Jiang, Y. G. (2023). AdaDiff: Adaptive step selection for fast diffusion models. arXiv. https://doi.org/10.48550/arXiv.2311.14768",
        "Zheng, Q., Chen, Y., Niu, Z., Ma, Z., Wang, X., Yu, K., & Chen, X. (2025). Accelerating flow matching based text to speech via empirical pruned step sampling. arXiv. https://doi.org/10.48550/arXiv.2505.19931",
    ]
    for ref in references:
        add_reference(doc, ref)

    props = doc.core_properties
    props.title = "Quality Constrained Policy Selection for Flow Matching Text to Speech Inference"
    props.subject = "AISTDS 2026 conference manuscript"
    props.author = "Anton Popov; [Academic Supervisor Name]"
    props.keywords = "text to speech, flow matching, adaptive inference, policy selection, WER, RTF"
    props.comments = "Prepared as a distinct expanded study with transparent citation of the related under review manuscript."
    doc.save(output)
    print(json.dumps({"output": str(output), "paragraphs": len(doc.paragraphs), "tables": len(doc.tables), "figures": 3}, ensure_ascii=False))


if __name__ == "__main__":
    args = parse_args()
    build(args.analysis_dir, args.output)
