from sprint.paper import DocInfo, PageInfo
from sprint.plan import (A3_DEST, A4_DEST, PLOT_DEST, Entry, Options, build_plan, mark_done,
                         reset_done)

A4P, A4L, A3P, A3L = (210, 297), (297, 210), (297, 420), (420, 297)
A1, A2, A0 = (841, 594), (594, 420), (1189, 841)


def entry(name, *sizes, error=None):
    e = Entry(name)
    doc = DocInfo(name, page_count=len(sizes), error=error)
    if not error:
        doc.pages = [PageInfo(i, w, h) for i, (w, h) in enumerate(sizes)]
    e.set_doc(doc)
    return e


def sample():
    return [entry("plan.pdf", A1, A1, A2), entry("bill.pdf", *[A4P] * 12),
            entry("details.pdf", A3L, A3P, A3L, A3L), entry("site.pdf", A0),
            entry("offer.pdf", A4P, A4L)]


def parts(plan_parts):
    return [(p.entry, p.pages) for p in plan_parts]


def test_routing_and_counts():
    plan = build_plan(sample(), Options(roll=610))
    assert (plan.a4.total, plan.a4.pending, plan.a4.included) == (14, 14, 14)
    assert (plan.a3.total, plan.a3.included, plan.a3_pages) == (4, 4, 4)
    assert sorted(plan.rolls) == [420, 610, 914]
    assert (plan.rolls[610].pending, plan.rolls[610].pending_mm) == (2, 1682)
    assert plan.rolls[420].lengths == {594: 1}
    assert plan.pending_rolls() == [420, 610, 914]
    assert parts(plan.a4_parts) == [(1, list(range(12))), (4, [0, 1])]
    assert parts(plan.a3_parts) == [(2, [0, 1, 2, 3])]
    # only the roll in the plotter is printed
    assert [(r.width_mm, r.length_mm, parts(r.parts)) for r in plan.plot_runs] == [(610, 841, [(0, [0, 1])])]
    assert plan.included == 14 + 4 + 2
    assert plan.entries[0].routes[2].dest == PLOT_DEST and plan.entries[0].routes[2].roll == 420


def test_runs_grouped_by_length_longest_first_across_files():
    entries = [entry("a.pdf", (594, 841), (610, 1000)), entry("b.pdf", (841, 594), (600, 1000.4))]
    plan = build_plan(entries, Options(roll=610))
    assert [(r.length_mm, parts(r.parts)) for r in plan.plot_runs] == [
        (1000, [(0, [1]), (1, [1])]), (841, [(0, [0]), (1, [0])])]
    assert plan.plot_included == 4
    assert plan.plot_included_mm == 2 * 1000 + 2 * 841


def test_a3_to_plotter_joins_the_297_roll():
    entries = [entry("a.pdf", A3P, A3L, (297, 630))]
    plan = build_plan(entries, Options(a3_to_plotter=True, roll=297))
    assert plan.a3.total == 0 and plan.a3_pages == 2
    assert plan.rolls[297].pending == 3
    assert [(r.length_mm, r.page_count) for r in plan.plot_runs] == [(630, 1), (420, 2)]
    # A3 switched off: not on the plotter either
    plan = build_plan(entries, Options(a3_to_plotter=True, a3_on=False, roll=297))
    assert plan.rolls[297].pending == 1 and plan.plot_included == 1


def test_switches_and_skipped_pages():
    entries = sample()
    entries[1].states[0].skip = True
    plan = build_plan(entries, Options(a4_on=False, plot_on=False, roll=610))
    assert (plan.a4.total, plan.a4.included) == (13, 0)
    assert plan.plot_runs == [] and plan.rolls[610].pending == 2
    assert plan.included == 4
    assert plan.entries[1].active == 11
    assert plan.entries[1].kinds["A4"] == 12


def test_unread_and_broken_files_are_left_out():
    entries = [Entry("reading.pdf"), entry("broken.pdf", error="not a PDF"), entry("a.pdf", A4P)]
    plan = build_plan(entries, Options())
    assert plan.entries[:2] == [None, None]
    assert parts(plan.a4_parts) == [(2, [0])]


def test_roll_by_roll_workflow():
    entries = sample()
    opts = Options(roll=610)
    plan = build_plan(entries, opts)
    mark_done(entries, plan.a4_parts + plan.a3_parts + [p for r in plan.plot_runs for p in r.parts])

    plan = build_plan(entries, opts)
    assert plan.included == 0                       # nothing left on this roll, printers done
    assert (plan.a4.pending, plan.rolls[610].pending, plan.rolls[610].total) == (0, 0, 2)
    assert plan.next_roll(610) == 420
    assert plan.entries[0].done == 2

    opts.roll = 420                                 # the user changed the roll
    plan = build_plan(entries, opts)
    assert plan.included == 1 and plan.plot_runs[0].length_mm == 594
    mark_done(entries, [p for r in plan.plot_runs for p in r.parts])
    assert build_plan(entries, opts).next_roll(420) == 914

    reset_done(entries, opts, PLOT_DEST, 610)        # print the 610 pages again
    plan = build_plan(entries, Options(roll=610))
    assert plan.rolls[610].pending == 2 and plan.rolls[420].pending == 0
    reset_done(entries, opts, A4_DEST)
    assert build_plan(entries, opts).a4.pending == 14
    assert build_plan(entries, opts).a3.pending == 0
    reset_done(entries, opts, A3_DEST)
    assert build_plan(entries, opts).a3.pending == 4


def test_a4_sheets_with_duplex():
    plan = build_plan(sample(), Options(duplex_a4=True))
    assert plan.a4_sheets == 6 + 1                  # per file: 12 pages -> 6, 2 pages -> 1


def test_oversize_pages_go_on_the_widest_roll():
    plan = build_plan([entry("big.pdf", (1000, 1500))], Options(roll=914))
    assert plan.oversize == 1
    assert plan.entries[0].routes[0].oversize
    assert [(r.width_mm, r.length_mm) for r in plan.plot_runs] == [(914, 1371)]
