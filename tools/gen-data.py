#!/usr/bin/env python3
"""Regenerate docs/data.js from SPI-Generator.py.

docs/data.js holds the FDEP pay item and task lists. The browser tool uses them
to work out which SOW Units cell a change order line belongs in. They are
generated from the Python app rather than typed out, so the two can never
disagree: when FDEP adds a pay item, edit the Python file and run this.

    python3 tools/gen-data.py           # rewrite docs/data.js
    python3 tools/gen-data.py --check   # exit 1 if docs/data.js is stale
"""
import ast
import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
SOURCE = ROOT / "SPI-Generator.py"
TARGET = ROOT / "docs" / "data.js"

# The generator writes SOW Units rows 11..502, columns G..R.
SOW_ROWS = 492
SOW_TASKS = 12


def function_list(tree, function, variable):
    """The literal assigned to `variable` inside `function`."""
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef) or node.name != function:
            continue
        for stmt in node.body:
            if (isinstance(stmt, ast.Assign) and len(stmt.targets) == 1
                    and isinstance(stmt.targets[0], ast.Name) and stmt.targets[0].id == variable):
                return ast.literal_eval(stmt.value)
    raise SystemExit("could not find `%s = [...]` inside %s() in %s"
                     % (variable, function, SOURCE.name))


def build():
    tree = ast.parse(SOURCE.read_text())
    pay_items = function_list(tree, "get_pay_items", "pay_items")
    tasks = function_list(tree, "get_tasks", "tasks")

    if len(pay_items) < SOW_ROWS:
        raise SystemExit("pay_items has %d entries; the SOW sheet needs %d"
                         % (len(pay_items), SOW_ROWS))
    if len(tasks) != SOW_TASKS:
        raise SystemExit("expected %d tasks in %s, found %d"
                         % (SOW_TASKS, SOURCE.name, len(tasks)))

    def literal(values):
        return json.dumps(values, indent=0).replace("\n", " ")

    return (
        "// Generated from ../../%s — do not edit by hand.\n"
        "// Regenerate with: python3 tools/gen-data.py\n"
        "var PAY_ITEMS = %s;\n"
        "var TASKS = %s;\n"
        % (SOURCE.name, literal(pay_items), literal(tasks))
    )


def main():
    text = build()
    if "--check" in sys.argv[1:]:
        if not TARGET.exists() or TARGET.read_text() != text:
            print("docs/data.js is out of date - run: python3 tools/gen-data.py",
                  file=sys.stderr)
            return 1
        print("docs/data.js matches %s" % SOURCE.name)
        return 0

    TARGET.parent.mkdir(parents=True, exist_ok=True)
    TARGET.write_text(text)
    print("wrote %s (%d bytes)" % (TARGET.relative_to(ROOT), len(text)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
