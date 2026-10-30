import os

import pytest

from sprint import jobs, printing


class FakeBackend(printing.Backend):
    def __init__(self, fail_on=()):
        self.calls = []
        self.fail_on = fail_on

    def print_file(self, printer, path, copies=1, paper=None):
        name = os.path.basename(path)
        if name in self.fail_on:
            raise printing.PrintError("printer jammed")
        self.calls.append((printer, name, copies) + ((paper,) if paper else ()))
        return True


@pytest.fixture
def ctx(tmp_path):
    messages = []
    context = jobs.JobContext(FakeBackend(), FakeBackend(), str(tmp_path),
                              lambda m, level="info": messages.append((level, m)),
                              printer_exists=lambda name: name != "Gone")
    context.messages = messages
    return context


def test_print_a4_sends_originals_and_subsets(ctx, make_pdf):
    only_a4 = make_pdf("text.pdf", [(210, 297)] * 2)
    mixed = make_pdf("mixed.pdf", [(210, 297), (420, 297)])
    no_a4 = make_pdf("drawing.pdf", [(841, 594)])
    sent = jobs.print_size(ctx, "A4", "Office", 2, [only_a4, mixed, no_a4])
    assert sent == 3
    assert ctx.backend.calls == [("Office", "text.pdf", 2, "A4"), ("Office", "mixed [A4].pdf", 2, "A4")]
    assert ctx.messages[-1][0] == "ok"


def test_print_reports_nothing_to_print(ctx, make_pdf):
    assert jobs.print_size(ctx, "A3", "Office", 1, [make_pdf("a.pdf", [(210, 297)])]) == 0
    assert ctx.messages[-1] == ("warn", "There are no A3 pages in the selected files.")


def test_print_checks_printer_and_files(ctx, make_pdf):
    path = make_pdf("a.pdf", [(210, 297)])
    assert jobs.print_size(ctx, "A4", "", 1, [path]) == 0
    assert jobs.print_size(ctx, "A4", "Gone", 1, [path]) == 0
    assert jobs.print_size(ctx, "A4", "Office", 1, []) == 0
    assert ctx.backend.calls == []
    assert [level for level, _ in ctx.messages] == ["warn", "error", "warn"]


def test_print_continues_after_a_failed_file(ctx, make_pdf):
    ctx.backend.fail_on = ("a.pdf",)
    paths = [make_pdf("a.pdf", [(210, 297)]), make_pdf("b.pdf", [(210, 297)])]
    assert jobs.print_size(ctx, "A4", "Office", 1, paths) == 1
    assert ("error", "  Error sending a.pdf: printer jammed") in ctx.messages


def test_plot_groups_by_length_and_collates_copies(ctx, make_pdf):
    a = make_pdf("a.pdf", [(594, 841)])
    b = make_pdf("b.pdf", [(210, 297), (841, 594), (600, 1000)])
    sent = jobs.plot(ctx, "610", "Plotter", 2, [a, b])
    assert sent == 3
    long, short = (610, 1000), (610, 841)
    assert ctx.plot_backend.calls == [
        ("Plotter", "b [610x1000].pdf", 1, long), ("Plotter", "b [610x1000].pdf", 1, long),
        ("Plotter", "a.pdf", 1, short), ("Plotter", "b [610x841].pdf", 1, short),
        ("Plotter", "a.pdf", 1, short), ("Plotter", "b [610x841].pdf", 1, short),
    ]
    assert ctx.backend.calls == []


def test_plot_requires_roll(ctx, make_pdf):
    assert jobs.plot(ctx, "", "Plotter", 1, [make_pdf("a.pdf", [(594, 841)])]) == 0
    assert ctx.messages == [("warn", "Select the roll width before plotting.")]


def test_plot_without_matching_pages(ctx, make_pdf):
    assert jobs.plot(ctx, "914", "Plotter", 1, [make_pdf("a.pdf", [(594, 841)])]) == 0
    assert ctx.messages[-1] == ("warn", "No pages for roll 914 mm.")
    assert ctx.plot_backend.calls == []


def test_plot_a3_sheets(ctx, make_pdf):
    assert jobs.plot(ctx, "A3", "Plotter", 1, [make_pdf("a.pdf", [(297, 420), (420, 297)])]) == 2
    assert ctx.plot_backend.calls == [("Plotter", "a.pdf", 1, (297, 420))]


def test_plot_reports_rejected_paper(ctx, make_pdf):
    ctx.plot_backend.fail_on = ("a.pdf",)
    assert jobs.plot(ctx, "610", "Plotter", 1, [make_pdf("a.pdf", [(594, 841)])]) == 0
    assert ("error", "  Error sending a.pdf: printer jammed") in ctx.messages


def test_cancel_stops_plotting(ctx, make_pdf):
    paths = [make_pdf(f"{n}.pdf", [(594, 841)]) for n in "abc"]

    class CancellingBackend(FakeBackend):
        def print_file(self, printer, path, copies=1, paper=None):
            ctx.cancel.set()
            return super().print_file(printer, path, copies, paper)
    ctx.plot_backend = CancellingBackend()
    assert jobs.plot(ctx, "610", "Plotter", 1, paths) == 1
    assert len(ctx.plot_backend.calls) == 1
    assert ctx.messages[-1] == ("warn", "Printing stopped.")
