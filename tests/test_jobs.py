import os

import pytest

from sprint import jobs, printing
from sprint.pdfdocs import read_document
from sprint.plan import A3_DEST, A4_DEST, PLOT_DEST, Entry, Options, build_plan, mark_done


class FakeBackend(printing.Backend):
    def __init__(self, fail_on=()):
        self.calls = []
        self.fail_on = fail_on

    def print_file(self, printer, path, copies=1, paper=None, duplex=None):
        name = os.path.basename(path)
        if name in self.fail_on:
            raise printing.PrintError("printer jammed")
        self.calls.append((printer, name, copies, paper, duplex))
        return True


@pytest.fixture
def ctx(tmp_path):
    messages, progress, sent = [], [], []
    context = jobs.JobContext(FakeBackend(), FakeBackend(), str(tmp_path),
                              lambda m, level="info": messages.append((level, m)),
                              printer_exists=lambda name: name != "Gone",
                              progress=lambda *args: progress.append(args),
                              sent=sent.append)
    context.messages, context.progress_calls, context.sent_parts = messages, progress, sent
    return context


@pytest.fixture
def entries(make_pdf):
    def load(*files):
        result = []
        for name, sizes in files:
            entry = Entry(make_pdf(name, sizes))
            entry.set_doc(read_document(entry.path))
            result.append(entry)
        return result
    return load


PRINTERS = jobs.Printers("Office", "Copier", "Plotter")


def test_prints_each_destination_with_its_paper(ctx, entries):
    files = entries(("text.pdf", [(210, 297)] * 2), ("mixed.pdf", [(210, 297), (420, 297), (841, 594)]),
                    ("drawing.pdf", [(841, 594), (594, 420)]))
    plan = build_plan(files, Options(roll=610))
    sent = [jobs.run(ctx, files, jobs.Job.from_plan(plan, dest, PRINTERS, copies=2, duplex=duplex))
            for dest, duplex in ((A4_DEST, True), (A3_DEST, True), (PLOT_DEST, None))]
    assert sent == [3, 1, 2]
    assert ctx.backend.calls == [
        ("Office", "text.pdf", 2, "A4", True),            # whole file: the original is printed
        ("Office", "mixed [A4].pdf", 2, "A4", True),
        ("Copier", "mixed [A3].pdf", 2, "A3", True),
    ]
    # only the 610 roll; the 420 roll page waits for the roll change
    assert ctx.plot_backend.calls == [
        ("Plotter", "mixed [610x841].pdf", 2, (610, 841), None),
        ("Plotter", "drawing [610x841].pdf", 2, (610, 841), None),
    ]
    assert ctx.progress_calls[-1] == ("plot", 2, 2)
    assert len(ctx.sent_parts) == 5 and ctx.problems == 0


def test_sent_parts_mark_pages_done(ctx, entries):
    files = entries(("a.pdf", [(210, 297), (841, 594)]))
    opts = Options(roll=610)
    plan = build_plan(files, opts)
    for dest in (A4_DEST, PLOT_DEST):
        jobs.run(ctx, files, jobs.Job.from_plan(plan, dest, PRINTERS))
    mark_done(files, ctx.sent_parts)
    plan = build_plan(files, opts)
    assert (plan.a4.pending, plan.plot_included) == (0, 0)


def test_failures_are_reported_and_not_marked(ctx, entries):
    files = entries(("good.pdf", [(210, 297)]), ("bad.pdf", [(210, 297)]))
    ctx.backend.fail_on = ("bad.pdf",)
    job = jobs.Job.from_plan(build_plan(files, Options()), A4_DEST, PRINTERS)
    assert jobs.run(ctx, files, job) == 1
    assert ctx.problems == 1
    assert [p.entry for p in ctx.sent_parts] == [0]
    assert ("error", "bad.pdf: printer jammed") in ctx.messages


def test_missing_printers(ctx, entries):
    files = entries(("a.pdf", [(210, 297), (420, 297)]))
    plan = build_plan(files, Options())
    for dest in (A4_DEST, A3_DEST):
        assert jobs.run(ctx, files, jobs.Job.from_plan(plan, dest, jobs.Printers("", "Gone", ""))) == 0
    assert ctx.problems == 2
    assert ctx.backend.calls == []


def test_cancel_stops_before_the_next_file(ctx, entries):
    files = entries(("a.pdf", [(210, 297)]), ("b.pdf", [(210, 297)]))
    job = jobs.Job.from_plan(build_plan(files, Options()), A4_DEST, PRINTERS)
    ctx.sent = lambda part: ctx.cancel.set()
    with pytest.raises(jobs.Cancelled):
        jobs.run(ctx, files, job)
    assert [call[1] for call in ctx.backend.calls] == ["a.pdf"]


def test_skipped_pages_are_left_out(ctx, entries):
    files = entries(("a.pdf", [(210, 297)] * 3))
    files[0].states[1].skip = True
    job = jobs.Job.from_plan(build_plan(files, Options()), A4_DEST, PRINTERS)
    jobs.run(ctx, files, job)
    assert ctx.backend.calls == [("Office", "a [A4].pdf", 1, "A4", None)]
    assert len(read_document(os.path.join(ctx.temp_dir, "a [A4].pdf")).pages) == 2
