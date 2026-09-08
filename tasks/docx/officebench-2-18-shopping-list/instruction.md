# Controlled shopping-list purchase brief

Prepare `/app/output/controlled_shopping_purchase_brief.docx` for a household purchase handoff. The five visible inputs are the only approved evidence:

- `baseline_shopping_list.xlsx` is the starting register. Preserve its source-row order for surviving items.
- `change_T01_reduce_apple.json`, `change_T02_remove_onion.json`, and `change_T03_add_yogurt.json` are approved change records. Apply them once, in ticket-ID order. Treat an item removed by a ticket as an audit entry, not a current purchase line; append an approved new item after surviving baseline items.
- `pricing_policy.json` supplies the costing rule. Apply its named-item exception and its “others” rate to every item in the final current list. For costing only, interpret each `amount` as pounds because the policy states prices per pound.

The document must be a reviewable working paper, not just a shopping table. Include:

- a clear control status and executive reconciliation summary;
- an editable current-purchase table, in the required order, with item, final amount, unit rate, extended cost, change origin, and source location;
- an editable three-ticket change ledger showing ticket ID, source file, action, before, after, signed delta, and disposition;
- an editable baseline-to-final reconciliation covering every baseline item plus every approved addition, including removed and unchanged items;
- calculation controls for baseline total amount, final total amount, signed amount change, current line count, total estimated cost, and average final amount per current line;
- a source register and concise evidence-boundary note. Do not invent a store, vendor, brand, tax, discount, purchase date, budget, package size, or price guarantee.

Make the evidence usable inside Word:

- use named styles and a native table-of-contents field;
- add bookmarks for the executive summary, current list, change ledger, reconciliation, and calculation notes, with working internal hyperlinks from the opening navigation area;
- anchor native Word comments to the changed apple line and the approved yogurt addition; each comment must identify the controlling ticket, source file, and enough source location detail to verify the change;
- put the detailed control tables in a landscape section, return the final notes to portrait, repeat table headers, prevent table rows from splitting, and use a status-bearing header plus PAGE/NUMPAGES footer fields;
- keep all data as editable text and native tables. Do not flatten the answer into screenshots, add macros/OLE objects, use external links, or include any non-DOCX output.

All calculations and conclusions must be traceable to the frozen inputs. Do not silently keep a removed item, apply a ticket twice, merge separate change events, or treat the old list as the current list.
