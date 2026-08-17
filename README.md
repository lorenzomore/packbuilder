# Packbuilder

A no-frills tool for keeping track of hiking and camping gear, and
building out what each person is actually carrying before a trip.

Everything lives in two places: a **gear library** (every item you
own or might pack) and **people & packs** (who's carrying what).

## Gear library

This is your master list — every piece of gear you might ever pack,
in one place, regardless of who ends up carrying it or on which trip.

- **Add an item** with a name, category (clothing, footwear, shelter,
  kitchen, and so on), an optional season, and a free-text note.
- **Season** marks something as Summer-only or Winter-only gear (a
  cap vs. a beanie, say). Leave it blank for anything that works
  year-round, like a water bottle or a camera.
- **Search and filter** by category to find things quickly once the
  list grows.
- **Edit or delete** any item with the pencil / trash icons on its
  row.
- **Upload a spreadsheet** to bulk-add gear instead of typing it in
  one at a time. Columns recognized: `Name`, `Category`, `Season`,
  `Notes`.
- **Download template** gives you a blank spreadsheet with dropdown
  menus already set up for Category and Season, so filling it in
  offline (or handing it to someone else to fill in) is fast and
  hard to get wrong. The category options live on a second sheet
  inside the file, so you can add or edit them freely — the dropdown
  updates automatically.
- **Clear all** wipes the entire library, if you want to start over.
  It asks you to confirm first.

## People & packs

This is where gear actually gets assigned.

- **Add a person** for each trip member.
- Each person can have **multiple packs** — useful for a day pack vs.
  an overnight pack, or separate summer/winter loadouts.
- **Add gear** opens a searchable list of everything in your library,
  grouped by category, with checkboxes so you can select several
  items at once and add them all together. There's also a Summer /
  Winter / All toggle to narrow the list down to what's actually
  relevant for the trip.
- **Add custom item** lets you throw something into a pack that
  isn't in your library and doesn't need to be — a borrowed item, a
  one-off, whatever doesn't deserve a permanent spot in the
  masterdata.
- Each item in a pack has its own **quantity**, a **packed**
  checkbox for last-minute checklist runs, and its own **notes**
  field — editing notes here only changes that copy, it never
  touches the library item.
- **Remove a person or a pack** at any time with the trash icon.

## Getting data in and out

- **Download template** / **Upload spreadsheet** (Gear library tab) —
  round-trip your gear list through Excel or Google Sheets.
- **Markdown export** (People & packs tab) — exports everyone's packs
  as one Markdown file, with checkboxes for every item. Built to drop
  straight into an Obsidian vault or any other Markdown-based notes
  app.

## A note on where your data lives

Packbuilder doesn't have a server or an account system — everything
you enter is saved locally in your browser. That means it'll still be
there next time you open the app in the same browser, but it won't
follow you to a different browser or device, and clearing your
browser data will clear it too.
