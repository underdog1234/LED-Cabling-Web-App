# LED Cabling Web App

Version `0.64.0`

Standalone React web app for planning LED wall layouts, signal port mapping, power outlet assignment, stock checks, deployment hardware, and PDF/settings/video exports.

## What It Does

- **Test Pattern Generator** (`generator.html`, also opened from the planner's toolbar) - a standalone page for any video output, projector, monitor or LED screen. See [Test Pattern Generator](#test-pattern-generator) below

- Build LED walls by rows and columns, or place panels freely (non-uniform layouts) with drag, edge-snap and joining
- Switch between `MG9` and `MT` panel profiles, plus `MG12` triangle and `MG13` curved variants
- Group panels into named **Sub-Screens**, each with its own chosen **colour**, and edit/patch each one in isolation - every other panel (including unassigned ones) is fully hidden and un-interactive while a sub-screen is active, with its name clearly shown; **All Screens** returns to the complete layout instantly, with no panel data ever altered
- Position the whole layout or individual sub-screens within a configurable-resolution **Output Canvas** (drag, numeric entry, align/snap tools, boundary/overlap warnings) for multi-processor / media-server mapping
- Select a **NovaStar processor model** (`VX1000 Pro` / `VX2000 Pro`), see its live pixel/port capacity vs. current usage, assign a video input per sub-screen or one input for the whole canvas, and generate a real, importable `.uprj` processor configuration file - validated against the selected processor's port count, per-port and total pixel limits, and canvas size, with a summary of what will be exported and any blocking errors or warnings before download
- Import projects from the Creative Layout Tool
- Patch signal and power manually or with auto-snake / automatic letter-patching routing, scoped to the active sub-screen when one is selected. The first panel of each signal chain shows its port number in a blue circle (top-left) and the first panel of each power chain shows its port number in an orange circle (top-right), in both the Panel Layout and the PDF Report. With **Do backup signal loop** enabled, the chain's last panel also shows the backup port number (the second half of the available signal ports, e.g. port 11 backs up port 1 on a 20-port setup) - the number of signal ports itself follows the selected NovaStar processor (10 for VX1000 Pro, 20 for VX2000 Pro, 20 if none is selected), and the backup half is hatched and unselectable in the Signal Patching panel
- See every signal and power run drawn over the panels, routed around the panel text so no label is ever covered. **Signal is always blue, power always orange**, and a hop both of them share is drawn as **one cable, with one arrow**, rather than a parallel pair. Every run carries a single outline `>` where it enters each panel. Identical on screen and in the PDF
- Hover any panel to read its **X / Y position** from the top-left corner of the layout, in mm and in content pixels
- Every panel carries its **row and column reference**, counted off the wall's own module grid - so a panel set half a module out reads as `row 5 & 6` instead of throwing the numbering of everything after it
- Flip the panel layout between `Back View` and `Front View`
- Both PDF Panel Layout pages draw each **sub-screen's boundary and name**, and **Generate PDF** can limit those pages to the sub-screens you pick - the page is then rebuilt around just those panels and says on its face that it shows part of the wall
- The PDF report carries an **Output Canvas page** - the full canvas resolution, every screen drawn to scale where it sits on it, and the same numbers as a table - plus a **key** on both Panel Layout views explaining every mark on the drawing
- A PDF **Weight breakdown** page: panels by type, every rigging and cable allowance with the arithmetic behind it, what is left out of the total, and a per-sub-screen panel weight
- Export a PDF report with portrait detail pages, a per-sub-screen summary page, plus both layout views in landscape - **Generate PDF** first asks which sections to include (everything ticked by default)
- Export native-resolution Test Pattern images - the whole wall and/or one per sub-screen, each at its own true output resolution - a full-screen canvas-only live Moving Test Pattern rendered pixel-for-pixel at its true output resolution (never scaled/stretched to fit the window - see [Pixel-Accurate Test Pattern](#pixel-accurate-test-pattern) below), or a downloadable looping WebM or MP4 video of it. The **MP4 is written to fixed delivery settings** - constant frame rate, High profile at a level the wall actually fits, half-second keyframes, no B-frames, Rec.709 limited range, exactly one loop long - and the download dialog lists them, with the frame rate and bitrate editable to match your show. Both the live view and the video ask which single surface to show first - the full canvas (where each sub-screen runs its own independent pattern) or one sub-screen on its own, at that screen's own resolution
- **Connectors are counted from the edges panels actually share** - one row per connector, per shared edge, with exposed edges adding nothing. Shape panels, corner panels and flat-laid corner panels each pull their own part, and a rotated panel's connectors move with its edges
- An **MG9 Corner panel can be laid in flat**: the same part off the same shelf, drawn without the corner hatch, needing the flat connector instead of the corner one
- Shaped panels (`MG12` triangle, `MG13` quarter circle) print the orientation code of the part they actually are - `LU` / `LD` / `RU` / `RD`, read from the front - so the drawing names the same stock line Stock Calculations counts. Left/Right is the corner the right angle is drawn in; Up/Down is the shelf's own vertical sense, which runs the opposite way (see `TRIANGLE_ORIENTATION`)
- Rotate panels by 45°, 90° or any custom angle, individually or as a multi-selected group (spacing/arrangement preserved); copy and paste panel groups with Ctrl/Cmd+C/V, with a cursor-following placement preview that snaps to the grid and nearby panels
- Toggle a vertical centre-line indicator on the Panel Layout (accounts for rotated panels' true outer bounds) with one control in **Overlays & displays** that governs both the on-screen layout and the PDF's Panel Layout pages
- Save and reopen settings as JSON (v5 format adds NovaStar processor/input selection; v3 sub-screens and output-canvas positioning, v2 free-panel and legacy grid formats still open)
- **Edit the stock list by hand** where the job needs it - type over any row's quantity, take a row off the list with an X, add any catalogue item the layout does not ask for, see at a glance how many rows are no longer what the tool calculated, and put every one of them back with one button. Edits are saved with the project and carried into the CSV, the PDF and the shortfall list, with the calculated figure shown alongside so nothing changes silently
- Check stock levels, shortfalls, and deployment hardware requirements, optionally checked against **Rentman** (see [Rentman Integration](#rentman-integration)) for live on-hand stock (reviewed before anything here is updated), what other projects have booked in the project's date range (the **peak on any one day** of it, not the sum of every job that touches it), and what is currently broken or under repair
- Collapse any section of the UI to reduce clutter on long projects

## Recent Changes In v0.55.0

**The Moving Test Pattern MP4 is a delivery file now**

It used to be whatever a constant-quality encode happened to produce, at the
live view's own 24fps. It is now written to fixed settings, and the download
dialog lists them before you commit to the encode:

| Setting | What the file gets |
| --- | --- |
| Format | MP4 / H.264 |
| Frame rate | **Constant** 60 fps by default - pick 24/25/30/50/60 to match your show |
| Resolution | The project's own content resolution |
| Profile / level | High / 4.2, or the lowest level that **can** carry the wall |
| Pixel format | 8-bit 4:2:0 (`yuv420p`) |
| Scan | Progressive |
| Keyframes | Every half second - 30 frames at 60fps, 15 at 30fps |
| B-frames | 0, for easier seeking |
| Bitrate | 8 Mbps target, 12 Mbps ceiling - both editable |
| Colour | Rec.709, limited (TV) range, tagged in the stream **and** in the file |

- **The frame rate is real, not a tag.** The pattern is now drawn at the rate
  the file is being made at, rather than the live view's gentler 24fps, and the
  output is forced to constant frame rate - a wall too big to redraw in time
  repeats a frame instead of the file going variable
- **The level follows the wall.** An LED wall is routinely wider than any
  broadcast format: 6,216 x 1,344 at 60fps needs level 5.2, and tagging that
  stream 4.2 would be a lie a hardware decoder is entitled to act on. 4.2 is
  used wherever it fits, and the dialog says when it does not
- **The colour conversion is spelled out** rather than inherited from whatever
  the browser tagged its recording as - which is the one thing a test pattern
  cannot afford to get wrong
- **The file is exactly one loop.** A 20-second recording came back 19.22
  seconds long, so the pattern jumped every time it repeated; the recording is
  now made longer than the loop on purpose and cut back to exactly one

Verified by reading the produced file's own boxes rather than trusting the
command line: `High / 4.2`, `colr nclx bt709/bt709/bt709` limited range, a
single frame-duration entry of **1200 frames at 60.000 fps**, 40 keyframes
exactly 30 apart, and no composition-offset table (so no B-frames).

## Recent Changes In v0.64.0

**A standalone Test Pattern Generator**

The planner's Output Canvas and Moving Test Pattern, opened up into their own
page at `generator.html` for any output - no LED panels needed. See
[Test Pattern Generator](#test-pattern-generator).

## Recent Changes In v0.63.0

**Every reference number is read from the front**

Row and column numbers were counted off the layout's own geometry, which is
the wall seen from the BACK. The test pattern already counted from the front,
so the panel the workspace called `C1` lit up as `37` on the wall - the two
numbers a crew actually compares were the two that disagreed.

- The **↓ row → column** labels, the PDF's chain tables (`R2 C37`) and the test
  pattern now all count columns **from the front**, the side you stand on to
  build the wall. One panel, one number, wherever you read it
- **Rows are untouched.** Flipping a wall left to right moves nothing up or
  down
- The number is fixed to the panel, so the Front/Back view toggle still never
  renumbers anything - it just changes which end you are looking at

**The hover readout is front-referenced too**

Hovering a panel gives its offset from the top-left of the layout **as seen
from the front**, in mm and in content pixels. That is the whole point of the
figure: it answers "which pixel of my content lands on this panel", and content
is authored for the front. Read off the back it named the mirror-image pixel -
the far end of the canvas - which is the one place a content op must not be
sent.

The front-left panel one row down now reads exactly what it should:

```
X 0 mm · 0 px
Y 500 mm · 168 px
from layout top-left
```

and the column number on that panel is `1`, so the label and the position can
no longer tell you different things.

## Recent Changes In v0.62.0

**Port bars show the port's own capacity, not your planning limit**

A signal port's load bar was drawn against the **Panels per Signal Port** box.
Typing a smaller number there painted a half-empty port red, and a bigger one
painted a full port green - the bar moved when nothing about the wall had.

- Each signal port now reads its **pixels against the 650,000 a port can
  drive**, exactly the way the power tiles already read amps against the 16A
  outlet. The px figure is printed under the panel count, so the bar can be
  checked rather than taken on trust
- It also counts a **mixed port** correctly: an MT panel is 16,384px against an
  MG9's 28,224, which no panel count can tell apart
- The **Panels per Signal Port** and **Panels per Power Outlet** boxes still do
  what they always did - they cap manual and auto patching. They just no longer
  decide what "full" looks like
- Power needed no change: its tiles and the phase bars were already drawn
  against the outlet and distro limits

**Shaped panel orientation codes**

`LU` / `LD` / `RU` / `RD` are read from the **front** of the wall. Left/Right is
the corner the right angle is drawn in; **Up/Down follows the shelf's own
vertical sense**, which runs the opposite way - a rotation-0 triangle draws its
right angle at the bottom and is the part named `LU`. The labels, the stock
lines and the pull sheet all go through one table, so they cannot disagree.

| Rotation | Triangle | Quarter circle |
| --- | --- | --- |
| 0 | `LU` | `RU` |
| 90 | `LD` | `LU` |
| 180 | `RD` | `LD` |
| 270 | `RU` | `RD` |

**Selections name the exact shaped part**

The selection chips split shaped panels by orientation - `2 x MG9 Triangle LU`,
`1 x MG9 Curved RD` - because each orientation is its own one-way part off its
own shelf, which is how Stock Calculations counts them. Four triangles facing
four ways are four different parts, and one line saying "4" was the wrong pull
every time.

## Recent Changes In v0.61.0

**Spare mesh modules on an MT wall**

| Code | Name | Quantity |
| --- | --- | --- |
| `12626` | YES TECH MT P3.9-7.8 LED Mesh Module | 3% of the MT panels, rounded up |

- Any MT panel on the layout puts the module on the list. It is **pure spare**:
  the row reads **Required 0**, with the 3% in the Spares column, because the
  modules go out fitted inside the panels - what travels separately is a
  handful to swap a dead one out on site
- Rounded **up**, so there is always at least one wherever there is any MT at
  all. A 192-panel wall takes **6**; 34 panels takes 2; 33 takes 1
- This is why MT panels themselves carry a **0% spare ratio** - nobody takes a
  spare mesh panel, they take the module that fails
- An MG9 or poster wall never sees the row

## Recent Changes In v0.60.0

**The selection says what it is made of**

Panel Tools showed "24 selected" and nothing else, so a selection of 20 standard
panels and 4 corners read the same as 24 plain ones - two different pulls, one
number.

- A chip per **physical part** now sits beside the count: `7 x MG9`,
  `1 x MG9 Triangle`, `2 x MT Corner`, biggest group first
- The split is the one **Stock Calculations** orders against, so a selection
  reads in the same terms as the pull sheet. The catalogue's own name for the
  part is on each chip's tooltip
- **Posters count as whole posters**, not the eight sections that selecting one
  pulls in
- A group holding inactive panels says so - `4 x MG9 (2 inactive)` - so Delete
  and Restore hold no surprises

**Copied panels open the workspace up**

The workspace stopped 300mm past the wall, and the pointer cannot reach past
the workspace, so there was nowhere to drop a copy except on top of what it
came from.

- With panels on the clipboard the workspace **opens out by the size of what is
  waiting to be pasted**, on all four sides, so a copy can be placed clear of
  the wall in any direction
- Per axis, so copying a wide, short strip does not open up a tall empty
  workspace it will never use
- It comes straight back off when the clipboard is cleared

**Sub-screen names on the test pattern**

- The **full-wall** test pattern PNG was the one export that showed no
  sub-screen names at all. It now outlines and names each screen in its own
  colour, the same way the moving pattern already did
- A **single sub-screen** rendered on its own now carries its name too. The
  boundary-and-name pass used to bail out whenever there was only one surface,
  which is exactly the case where the name matters most - the shape alone no
  longer says which screen you are looking at
- A plain wall with no sub-screens still has nothing drawn over it

## Recent Changes In v0.59.0

**A sub-screen's pixel count is the rectangle it stands in**

A sub-screen with a gap or a step in it was quoted one module column short on
the Output Canvas and in the PDF, while the test pattern rendered the right
figure. On the file this came from, the Left screen read **3,528 x 1,344** in
both places and **3,696 x 1,344** on the test pattern - a 168px column, one
whole panel, missing from the number the content gets built to.

- A sub-screen's resolution is now its **footprint**, the same way the wall's
  own Resolution already was. The two agree exactly on a solid rectangle and
  part company the moment the shape is stepped, which is where the footprint
  is the right answer
- The **Complete layout resolution** line and the whole-layout entry had the
  same fault when a project has no sub-screens: 5,880 where the wall is 6,216
- This is also the only figure consistent with where the panels are actually
  drawn on the canvas, which has always used their physical mm offsets - a
  sub-screen's own panels could sit outside the box the canvas view drew
  around them
- The **packed** space is unchanged and still right where it is used: the
  NovaStar cabinet topology in the `.uprj` export, where a processor has no
  such thing as an empty gap pixel

**Vertical cable runs sit on the panel, not on its edge**

A vertical run was drawn about 0.9px from a panel's left edge, close enough
that its outline actually overhung the boundary and the run read as if it
belonged to the join rather than the panel.

- Both lanes now sit **12% of the way in**, the same inset the horizontal lanes
  already used - the signal run moves from 0.9px to 7.1px off the edge on a
  full panel, roughly eight times further in, beside the text rather than on
  the border
- **No label got smaller.** A full panel still prints at 10px in the PDF and
  9px on screen, and a poster section still gets its 6px
- That is paid for by centring the label block on the space **beside** the
  lanes instead of on the panel itself. Centring on the panel meant reserving
  the same band twice - once on the left where the runs are, and once on the
  right purely to stay symmetrical - and it was that second, empty reservation
  that used to decide how big the text could be
- The rule that a run never crosses a panel's text still holds, and is now
  tested against where the text really sits

## Recent Changes In v0.58.0

**Cancelled jobs no longer hold stock**

Rentman keeps a cancelled job's equipment lines rather than deleting them, so
they kept coming back from the date-overlap query and were counted like any
other booking. A job called off months ago could still show its panels as out,
and the tool reported a shortage over kit sitting on the shelf.

- A booking whose job is **cancelled is never counted** - not in **Out on the
  day**, not in the sum beside it, not in **Available Stock**
- It is still **listed**, marked *cancelled, not counted*, with a line saying
  what it would have added: *"200 on 1 cancelled job is not counted - that kit
  is on the shelf."* A planner can see the booking in Rentman, so silently
  dropping it would raise the same question the old behaviour did
- The PDF's Other Projects & Repairs page says the same, and its intro now
  spells it out
- Matching is on the status **stem**, because Rentman spells it two ways: the
  API's status name is the US "Canceled" while the UI shows "Cancelled".
  Checked against this account's real status list - Pending, Canceled,
  Confirmed, Prepped, On location, Returned, Inquiry, Concept
- **Only cancelled is dropped.** A Pending or Inquiry job is kit somebody is
  seriously expecting to send out, and quietly freeing it would be the same
  mistake in the other direction
- The Worker has the same rule, so its own `peakRequired` and `remaining`
  agree - but the app works this out from the bookings it is given, so it is
  right **without redeploying**

**Levelling packers on anything that stands on the floor**

| Code | Name | When |
| --- | --- | --- |
| `12374` | Assorted Levelling Packers & Shims | Every deployment except **Flown** |

One assorted set, which is how Rentman stocks it (quantity 1) - not a
per-panel count. A flown wall hangs off the bar and has nothing to pack;
ground support, floor and no-support walls all sit on whatever the venue's
floor happens to be doing. With no deployment type chosen the row stays off
the list, same as every other deployment-driven item.

**Also:** the VX2000 Pro's shelf quantity is now `3`, confirmed against
Rentman along with both processor codes and the packers' - all four read back
exactly as entered.

## Recent Changes In v0.57.0

**The processor is now on the equipment list**

Picking a Processor Model puts that processor on the stock list, 1 per project:

| Processor Model | Code | Name |
| --- | --- | --- |
| NovaStar VX1000 Pro | `12247` | NovaStar VX1000 Pro LED Processor |
| NovaStar VX2000 Pro | `12353` | NovaStar VX2000 Pro LED Processor Rack |

- The row **follows the selection**, so changing model swaps the line rather
  than leaving two processors on the sheet, and **None selected** (manual port
  planning) leaves the processor off the list entirely
- Quantity is **1** - deliberately not one per 13.1MP. The Processor Capacity
  readout and the NovaStar validation already say when a wall is too big for
  the model picked, and the answer to that is a bigger processor, not a second
  unit appearing quietly on the pull sheet
- Both are catalogue items, so they reach the on-screen table, the CSV, the PDF
  stock table and the Rentman stock check like anything else - and either can
  be added by hand from **Add an item** when the unit going out is not the one
  the project was planned on
- Shelf quantities are the MG9 catalogue's own `vx1000`/`vx2000` figures (2
  each), read from there rather than retyped, so a stock check updates one
  number

**Also: every catalogue code is now tested.** No two items may sit on one code
(a duplicate reads as a doubled requirement and breaks the Rentman comparison),
every shelf quantity has to be a whole number, and the codes that read
backwards on purpose - triangle `12399` above quarter-circle `12398`, the floor
parts on `12250`/`12251`/`12252` - are pinned, so a future tidy-up cannot swap
them back.

## Recent Changes In v0.56.0

**Fixed: availability was adding up jobs that never clash**

The Rentman availability check asked "what is booked in this window" and added
it all together. Two jobs of 100 panels, one on the 4th and one on the 8th,
came back as **200 unavailable** - though the first lot is on the shelf again
days before the second goes out. On a busy month that reported shortages that
did not exist.

- It now asks the question that matters: **the most of an item out on any ONE
  day of your range**. Two jobs of 100 on different days are 100 unavailable;
  two that share a day are 200
- The column is **Out on the day**, and opening it names the peak day, the
  bookings that make it up, and - when the two differ - says what the old sum
  would have been and why it is not the answer: *"260 is booked across the
  whole range, but these jobs do not all run together - only 160 is ever out on
  one day"*
- The PDF's Other Projects & Repairs page says the same, marking the bookings
  in the peak with a `*`
- **Days, not hours.** Two jobs sharing a day are competing for the same panels
  even when one derigs in the morning and the other rigs in the afternoon, and
  a planner wants that flagged rather than smoothed away by a clock
- A booking only counts for the days it **shares with your range** - a
  month-long sub-rental running either side of your job takes up your days, not
  its own
- The app works this out from the bookings the proxy already returns, so it is
  right **without redeploying the Worker**. The Worker has been fixed too - it
  now returns `peakRequired` alongside the old sum, and its `remaining` uses
  the peak - but nothing waits on that going live

## Connector Rules

Connectors are counted from the edges panels **actually share**, found from their
real rotated positions. Each shared edge is counted **once**, never once per
panel, and an edge with nothing on the other side adds nothing.

A **horizontal edge** is a horizontal line, so one panel sits on top of the
other; a **vertical edge** is a vertical line, so they sit side by side.

| Panels joined | Edge | Connector | Per edge |
| --- | --- | --- | --- |
| Plain MG9 + Plain MG9 | Stacked (horizontal edge) | **none** | - |
| Plain MG9 + Plain MG9 | Side by side (vertical edge) | **none** | - |
| Plain MG9 + MG9 Corner (folded) | Stacked (horizontal edge) | 150 Connector | 3 |
| Plain MG9 + MG9 Corner (folded) | Side by side (vertical edge) | 150 Connector | 3 |
| Plain MG9 + MG9 Corner (laid flat) | Stacked (horizontal edge) | 150 Connector | 3 |
| Plain MG9 + MG9 Corner (laid flat) | Side by side (vertical edge) | 150 Connector | 3 |
| Plain MG9 + MG12 triangle / MG13 curve | Stacked (horizontal edge) | 150 Connector | 3 |
| Plain MG9 + MG12 triangle / MG13 curve | Side by side (vertical edge) | Horizontal Connector | 2 |
| MG9 Corner (folded) + MG9 Corner (folded) | Stacked (horizontal edge) | MG9 Corner Connector | 3 |
| MG9 Corner (folded) + MG9 Corner (folded) | Side by side (vertical edge) | MG9 Corner Connector | 3 |
| MG9 Corner (folded) + MG9 Corner (laid flat) | Stacked (horizontal edge) | MG9 Corner Connector | 3 |
| MG9 Corner (folded) + MG9 Corner (laid flat) | Side by side (vertical edge) | MG9 Corner Connector | 3 |
| MG9 Corner (folded) + MG12 triangle / MG13 curve | Stacked (horizontal edge) | 180 Connector | 3 |
| MG9 Corner (folded) + MG12 triangle / MG13 curve | Side by side (vertical edge) | 150 Connector | 2 |
| MG9 Corner (laid flat) + MG9 Corner (laid flat) | Stacked (horizontal edge) | 150 Connector | 3 |
| MG9 Corner (laid flat) + MG9 Corner (laid flat) | Side by side (vertical edge) | 150 Connector | 3 |
| MG9 Corner (laid flat) + MG12 triangle / MG13 curve | Stacked (horizontal edge) | 180 Connector | 3 |
| MG9 Corner (laid flat) + MG12 triangle / MG13 curve | Side by side (vertical edge) | 150 Connector | 2 |
| MG12 triangle / MG13 curve + MG12 triangle / MG13 curve | Stacked (horizontal edge) | 180 Connector | 3 |
| MG12 triangle / MG13 curve + MG12 triangle / MG13 curve | Side by side (vertical edge) | 180 Connector | 2 |

The order of the pair never changes the answer. A corner panel meeting a shape
has its own rule - 3 x 180 stacked, 2 x 150 side by side - whether it is folded
round the corner or laid flat.

**MT panels** are not in this table: an MT join needs no connector. An MT panel
can be marked as a **corner** (the same panel off the same shelf, so the panel
count does not change), and each MT corner adds **2 x MT Corner Connecting
Bracket and 8 x Bolt** - counted per corner panel, not per joined edge.

### What this does not count, and why

Each of these is pinned as a test in `src/connectors.test.ts`:

- **Plain MG9 to plain MG9 adds nothing.** MG9 panels ship with their own
  vertical connectors and three horizontals, so a plain panel-to-panel join
  needs nothing pulled from the shelf
- **MG9 Vertical Connector (12480)** is never chosen by a join rule for the
  same reason - it comes only from the Modular Frame Bottom Beam allowance, 4
  per 1m beam
- **A brick bond is not seen as a join.** Anchors sit every 100mm along a 500mm
  edge, so a neighbour offset 100mm or 200mm still lines up - but one offset
  **half a panel (250mm)** lines up with nothing, and the pair reads as
  unjoined. On a plain MG9 wall this costs nothing, since such a join needs no
  connector anyway; it only under-counts where a corner or shaped panel sits on
  a staggered join
- **Panels turned to an odd angle** (45 degrees, say) are not seen as joined.
  Those joins are made with hardware outside this catalogue
- **A triangle never joins along its hypotenuse**, nor a quarter circle along
  its arc - correctly, since there is nothing there to bolt to


## Recent Changes In v0.54.0

**Corner panels meeting a shape have their own rule**

- A corner panel against an MG12 triangle or MG13 quarter circle now takes
  **3 x 180 Connector stacked** and **2 x 150 Connector side by side**, whether
  the corner is folded round the corner or laid flat. It used to count as a
  plain MG9 there, which pulled the wrong part for both orientations
- A **plain** MG9 against a shape is unchanged: 3 x 150 stacked, 2 x Horizontal
  Connector side by side

**MT panels can be marked as corners**

- The panel variant picker now offers each panel type only what it physically
  has: every shape for MG9, **Standard or Corner for MT**
- An MT corner is the **same MT panel off the same shelf**, so the panel count
  does not change. What it adds is hardware: **2 x MT Corner Connecting Bracket
  (12274) and 8 x Bolt (12275) per corner panel**, counted per panel rather
  than per joined edge, which is how the part is fitted
- MT joins otherwise need no connector, which is why MT is not in the connector
  table

**The rest of the audit, settled**

- Side-by-side corner joins stay at **3 per edge** - confirmed, not an oversight
- Plain MG9 to plain MG9 stays at **nothing**: the panels ship with their own
  vertical connectors and three horizontals
- The MG9 Vertical Connector comes in the panels too, so no join rule asks for one
- Odd-angle joins are made with hardware outside this catalogue and are not counted
- All of it is written down under [Connector Rules](#connector-rules) and pinned in the tests

## Recent Changes In v0.53.0

**Shape to shape is always the 180 Connector**

- Two shaped panels (MG12 triangle, MG13 quarter circle) meeting leg to leg now take the **180 Connector whichever way the edge runs**. Stacked was already 3 x 180; side by side took the *150 Connector* and now takes **2 x 180**

**The whole connector table, written down**

- Every pair of panel kinds and both edge orientations - 20 rows - are listed under [Connector Rules](#connector-rules) below, and pinned row by row in `src/connectors.test.ts`, so the table and the code cannot drift apart
- The **limits of the model** are pinned there too, as tests: what the edge finder does not see is now recorded rather than assumed (see the list under the table)

## Recent Changes In v0.52.1

**Panel Layout pages print sharply at A3**

- The layout drawing is a picture, and it was rendered at **300 DPI for the A4 page it sits on**. Print that A4 page enlarged onto A3 and the same pixels are stretched 1.41x - 212 DPI - and the panel text, which is only about a millimetre tall to begin with, went to pieces
- It now renders at **600 DPI**, so an A3 print lands at 424 DPI and even A2 still has 300 to work with
- The cost is the file: a 141-panel wall goes from about 0.9MB to 2.0MB, and the report takes roughly ten seconds longer to build. Everything else in the PDF is vector and was always sharp at any size
- A ceiling on the image size guards the canvas: the print area is fixed, so at 600 DPI nothing can exceed about 23 megapixels, and a future DPI rise cannot quietly hand the browser a canvas too big to allocate - which fails by returning a blank image rather than by raising an error

## Recent Changes In v0.52.0

**Fixed: the height ruler started in the wrong place**

- The metre marks up the side of the wall were laid out from the **top** and merely relabelled, so on a wall that is not a whole number of metres high - three 0.5m panels being the plain case - **`0m` landed between two panels instead of on the ground**
- The ruler is now measured **up from the bottom of the wall**: `0m` is the bottom edge, always, and the whole-metre grid lines move with it so the numbers sit on the lines they belong to
- Fixed in the Panel Layout and in the PDF's layout pages alike

**A weight breakdown page in the PDF**

- New **Weight breakdown** section in Generate PDF: panels type by type (count x each = weight), then every rigging and cable allowance with the sum that produced it - `12 x 1.9kg (MG9) + 0 x 5.9kg (MT)`, `3 outlets in use x 3kg` - and the grand total
- Allowances that are switched **off** are listed too, greyed, with "No" in the "In the total" column. An allowance nobody can see is one nobody can question
- With sub-screens, a per-screen panel-weight table follows. Panels only: the rigging allowances are worked out across the whole wall, and splitting them per screen would be inventing a number

## Recent Changes In v0.51.0

**Sub-screens on the PDF's Panel Layout pages**

- Both Panel Layout pages now draw each **sub-screen's boundary** the way the workspace does: a dashed box just outside that screen's panels, in the screen's own colour, with its name above it. The key gained a row explaining the mark
- **Generate PDF** now lists every sub-screen. Untick one and its panels are left off those pages, which are then built around what is left - its own bounding box, its own rulers, its own centre line - exactly as the workspace does when a sub-screen is opened for editing. Panels in no sub-screen are their own entry
- A page showing only part of the wall says so on its face: *"Showing Main Screen, Side Right only - 38 of 46 panels. The figures below are the whole wall's."*
- A cable run to a panel the page does not show is left out whole, rather than drawn heading off into white space
- Everything stays ticked by default, so a report nobody touches is the same report as before

**Metre marks moved out of the way**

- In the Panel Layout and in the PDF, the metre numbers now sit at the **edge of the canvas** rather than tight against the wall. That band belongs to the sub-screen name labels, and the two were landing on top of each other. The grid lines and ticks still line up with the numbers, so nothing is lost by moving them out

**Moving test pattern**

- The **black outline around the wall info text is gone** - at a distance it read as a smear around every letter
- The info text now lands **on panels**. On a stepped or ring-shaped wall the middle of the bounding box has no LED in it, and the text hung in the air where it could not be read at all; it now moves to the nearest spot that is actually covered
- **Panel numbers are much bigger on MT.** Their size came from the panel's short side, so a 256 x 64 MT panel carried 6px text - unreadable on the wall it was made for, with the 256px of width beside it unused. The size now has a pixel floor and is capped by what the two-line block can fit, which leaves MG9 close to where it was and roughly triples MT

## Recent Changes In v0.50.0

**Items can be added to the stock list, not just edited**

- A new **Add an item** control under the Stock Calculations table puts any item in the catalogue on this project's list, with whatever quantity you type
- Added rows are marked **added by hand** on the list and carry a `*` in the PDF, which prints `Put on this list by hand - nothing on this wall asks for it` underneath with the item and quantity - the same treatment an edited quantity gets
- They go into the CSV, the shortfall list and the Rentman stock check like any other row, and are saved with the project
- Only the item's **code** is stored, so its name and shelf quantity are read from the catalogue each time the list is built rather than frozen into the project file
- An item already on the list is not offered, so the same code can never appear twice

**Three items added to the catalogue**

| Code | Item | Stock |
| --- | --- | --- |
| `12272` | YES TECH Patch F/M - F/M Signal Cable | 14 |
| `12274` | YES TECH MT Corner Connecting Bracket | 100 |
| `12275` | YES TECH MT Corner Connecting Bracket Bolt | 400 |

- These are the items whose codes this catalogue had against its own floor parts until v0.49.1. **Nothing works out a requirement for them** - the deployment-hardware formulas cover MG9 only, and the signal count has its own joiner and joiner cable - so they never appear on a list by themselves. They are in the catalogue so they can be added to one by hand, and so a stock check knows the codes
- **32A 3Ph PDL Power Adaptor (`6650`) stock is now 10**, so a project needing one no longer reads as a shortfall

## Recent Changes In v0.49.1

**Stock codes corrected**

- **Triangle is `12399`, quarter circle is `12398`** - the opposite way round to what the names suggest, and the opposite way round to what this catalogue had. A stock check on the old codes was returning each other's item
- Three floor items were on codes that belong to something else entirely. They now read:

| Item | Was | Now |
| --- | --- | --- |
| 500mm x 500mm Tempered Glass Floor Cover | `12272` | `12250` |
| Modular Frame Floor Reinforcement Bar | `12274` | `12251` |
| Modular Frame Floor Taper Mounting Pin | `12275` | `12252` |

- `12272`, `12274` and `12275` are the **Patch F/M - F/M Signal Cable**, the **MT Corner Connecting Bracket** and its **Bolt** - none of which this app calculates a requirement for, so they are not in the catalogue
- Their shelf quantities go back to the catalogue's own figures (`384`, `384`, `1,536`): the `14`, `100` and `400` that came back in the stock check were counts of those other three items, not of these
- If a **Get Current Stock** result was applied before this, its confirmed figures are stored against the old codes and no longer match anything - run the check again to refresh them

## Recent Changes In v0.49.0

**The stock list can be edited by hand**

- Every row of Stock Calculations now has its **quantity in a box you can type in**, and an **X** to take the item off this project's list altogether
- A **Reset to calculated** button at the top of the card puts every one of them back, and a chip beside it says how many rows are currently set by hand - so a list somebody has been editing can never be mistaken for the tool's own answer
- An edited row shows the figure the tool worked out beside the new one (`edited - tool says 10`), and the Required / Spares columns still show where that figure came from
- Removed rows are listed under the table with a **Put back** button, so taking one off is never a one-way door
- Edits reach **everything**: the CSV, the PDF Stock Summary, the shortfall list. The PDF marks each edited row with `*` and prints, under the table, exactly what was changed and what was taken off - a printed pull sheet never differs silently from what the tool calculated
- They are saved **with the project** (`formatVersion` 7), because they are decisions about this job. Rentman's confirmed stock figures still live per browser, because those are facts about the warehouse

**Panel text stays on the panel in the workspace too**

- v0.48.0 moved the text inside the shape for the PNG exports and the PDF. The on-screen Panel Layout now uses the **same rule**: the foot of the panel first, because that is the band the cable router keeps clear, and a point inside the silhouette only when that foot is the shape's empty corner

**Stock figures updated**

- Shelf quantities updated from a Rentman stock check: 150 Flat Connector `320`, 180 Connector `80`, Horizontal Connector `1,170`, Dance Floor Feet `576`, Floor Reinforcement Bar `100`, Floor Taper Pin `400`, Tempered Glass Floor Cover `14`, MG9 panels `320`, 15m PowerCON `93`, 15m Signal `57`
- The distros and the two 15m cables are one stock line each for the business, not one per panel type - they used to read `0` on an MT wall, reporting every cable as a full shortfall
- Five codes in that check came back under a different name than this catalogue had for them. They were the catalogue's codes that were wrong, and they are corrected in v0.49.1 below

## Recent Changes In v0.48.0

**Row and column numbers a stepped wall can actually use**

- Row and column references are now read off the **wall's own module grid** rather than by grouping panels that happen to line up. A panel dropped half a module out used to be given a row of its own, so a three-high wall counted as five rows and **every number after that panel was wrong**
- A panel that straddles two rows now says so: **`row 5 & 6`**, the way a crew would say it on site. A panel spanning more than two reads as a range (`3-5`), and every panel that does sit on the grid keeps the number it always had
- The grid's cell size follows what the wall is mostly built from, so a wall of 1m-wide `MT` panels still numbers them 1, 2, 3 - not 1&2, 3&4
- The same numbers now appear everywhere they are read: on the panel in the workspace, in the PNG and moving test patterns, on the PDF layout pages, and in the signal/power chain columns (`R1 C1 -> R3 C1`)
- The reported wall size follows them too - **"Panels: 141 active across 8 rows x 37 columns"** - so the header and the panel labels can no longer disagree, and it matches the Active support span beside it

**Panel text stays on the panel**

- On a **triangle** or **quarter circle**, one corner of the panel's rectangle is empty. The label was drawn there regardless, so in the PNG exports it fell outside the shape - cut off with it, or sitting over the panel next door
- Shaped panels now place their text **inside the silhouette**, on a point that turns with the panel, so it lands on the lit area at every rotation and in both front and back views
- In the PDF, the text still starts at the **foot of the panel**, where the cable router already leaves it room; it only moves inward - shrinking to fit first - when the foot of that shape has nothing to sit on, such as a triangle standing on its point
- Rectangular and corner panels are untouched

## Recent Changes In v0.47.0

**Fixed: the PDF could fail outright**

- The 32A distro adaptor added in v0.46.0 carries a **Greek phi** in its name (`3Φ`). jsPDF's built-in fonts only cover WinAnsi, so that one character made the line print as `3 2 A  3 |  P D L ...` - spaced-out nonsense - and pushed jsPDF onto a wide-encoding path that can fail the whole export
- Every string in the report now passes through one guard on its way into the document: `3Φ` prints as `3Ph`, arrows and dashes become their ASCII equivalents, and anything left that the font cannot set is dropped rather than corrupting the line it is in. **Nothing had to be remembered at each of the hundred places the report writes text**
- When a PDF does fail, the message now **names the error** instead of saying "check console" - so the one person who can report the fault has something to report

**Output Canvas in the PDF**

- A new **Output Canvas** page: the **full canvas resolution**, the canvas drawn as a frame with **every screen to scale where it has been placed on it**, and the same figures as a table - canvas X, canvas Y, resolution, right/bottom edge and panel count per screen
- Each screen is outlined in its own identity colour, the same colour it carries in the workspace
- Drawn as **outlines on white, no solid fills** - it is a page to print, and a page of solid dark is a page of ink
- Anything that will not map is called out underneath: a screen hanging off the canvas, a negative position, or two screens sharing pixels
- Ticked by default in **Generate PDF**, and untickable like every other section

**A key on the Panel Layout pages**

- Both the Back View and Front View pages now carry a **key** in the top-right corner, where the header already left the page empty - so the drawing keeps its full size
- It covers the signal line, the power line, a run carrying both, the `>` direction mark, the centre line, the signal and power chain-start badges, and how to read `1 > 2`, `P1 (3)` and the `LU/LD/RU/RD` shape codes
- Drawn from the same colours and the same chevron the layout itself uses, so the key cannot drift from what it is explaining

## Recent Changes In v0.46.0

**Connectors are now worked out from the layout**

- Connector quantities come from the **edges panels actually share**, found from panel positions, so free-form layouts and rotated panels are handled the same as a plain grid. Each shared edge is counted **once**, not once per panel, and an exposed edge - one with nothing on the other side - adds nothing
- A **horizontal edge** is a horizontal line, so one panel sits on top of the other; a **vertical edge** is a vertical line, so they sit side by side. It is read off the panels' real rotated positions, so **turning a panel turns the edges it joins along** and its connectors change with it
- Horizontal edges, 3 per edge: corner-to-corner **both flat** takes the *150 Connector*, corner-to-corner **otherwise** the *MG9 Corner Connector*, shape-to-MG9 the *150 Connector*, shape-to-shape the *180 Connector*
- Vertical edges, 2 per edge: shape-to-MG9 takes the *Horizontal Connector*, shape-to-shape the *150 Connector*
- Everything not in those rules keeps the behaviour it had: corner-to-plain-MG9 is still 3 × *150 Connector*, and two plain MG9 panels still need nothing
- Several rules share one stock item, so each item gets **one row** with the rules that asked for it spelled out in its method text - two rows on one code would have read as a duplicate requirement

> **Two codes, not one.** The brief gave both the *150 Connector* and the *MG9 Corner Connector* code **12260**. The catalogue has them as separate items - 12260 for the 150 Connector and **12258** for the MG9 Corner Connector, each with its own shelf quantity - so they are kept apart. Pulling one code for both rules would have ordered the wrong part for half of them.

**MG9 Corner panels can be used flat**

- **MG9 LED Corner Panel (flat)** is a new choice wherever the panel variant is set. It is the **same physical part** - same stock line, same shelf, same spare - laid in flat instead of folded round a corner
- Drawn without the corner hatch, so which panels are actually turning a corner reads at a glance
- A corner-to-corner join counts as **flat only when both panels are set flat**; one panel still folded means the corner part is what fits

**Accessories that come with the equipment**

- A **32A distro** brings **1 × 32A 3Φ PDL - 32A 3Φ Ceeform Power Adaptor** (`6650`) per distro
- A **Modular Frame Bottom Beam 1m** brings **4 × MG9 Vertical Connector** (`12480`) per beam
- Both move with the equipment they hang off, and both appear in the stock totals and the CSV export like any other line

> Three of the new items - the *180 Connector* (12476), the *Horizontal Connector* (12623), the *MG9 Vertical Connector* (12480) and the *distro adaptor* (6650) - have no shelf quantity in this catalogue yet, so they start at **0** and read as a full shortfall until Rentman fills them in or an override is set.

**PDF labels are easier to read**

- Panel labels, metre markings and the view heading in the PDF's Panel Layout pages now carry a **thin white casing behind the text**. The fill colour is unchanged and the outline is stroked underneath it, so characters stay sharp - it just lifts them off the panel fill and the cabling crossing it

## Recent Changes In v0.45.0

**A cable that doubles back on itself now shows as two cables**

- Where a chain turns back on itself, or a return leg retraces one that went out earlier, the two runs landed on the same lane. That drew as a **single line with two arrows piled on it** - it looked like one cable
- The second run now **steps clear onto its own lane and keeps its own arrow**, so both read. It keeps its true end points and moves only in between, so it still joins the hops either side of it. The step is always **away** from the panel's text, so it can never cost the clearance the label size is worked out against

**Wall Resolution was under-reporting a stepped layout**

- The wall's **Resolution** was taken from the longest single row of panels. On a stepped or L-shaped wall that is narrower than the wall itself: the 18.5m test layout reported `5880 × 1344` when it stands across all 37 module columns - `6216 × 1344`
- Resolution, aspect ratio, reduced ratio, recommended content resolution and best standard output are now measured from the wall's **physical footprint** at its finest pixel pitch. **A rectangular wall is unchanged** - the two figures agree exactly there
- The NovaStar `.uprj` export is deliberately **not** affected: a processor's cabinet topology has no such thing as an empty gap pixel, so it keeps its own packed space (see `canvasModel.ts`)

**Shaped panels name the part they are**

- An `MG12` triangle or `MG13` quarter circle is a one-way piece: where its rotation puts the right-angle corner decides which physical part it is. Each panel now prints that code - **`LU`, `LD`, `RU` or `RD`, read from the front** - next to its shape symbol, instead of a generic "rotated" marker
- Same four buckets Stock Calculations already counts against the shelf, so the drawing and the stock list name the same thing

## Recent Changes In v0.44.0

**Cabling is far easier to follow on a complicated layout**

- **Signal is always blue and power always orange**, whatever port they belong to. The port is already named by the panel fill and by the numbered badge on the chain's first panel, so one colour per service is worth more than twelve colours per port once a wall gets busy
- **Where signal and power run between the same two panels, they are now drawn as ONE cable, with ONE arrow** - a blue run with the power orange dashed over it - instead of two lines side by side. On a wall patched with **Match Power To Signal Pattern** that is almost all of the cabling (about nine out of ten hops on the 141-panel test layout), so the drawing has roughly half as many lines on it
- Where the two genuinely do part company they stay separate, side by side, exactly as before
- **Signal runs now carry direction arrows too** - the same single outline `>` where the run enters each panel that power already had. One per panel entered, one for a shared run, and never repeated along a run
- **Runs no longer zigzag.** Each panel now has one anchor point where the two clear bands of the panel meet, so a hop is a straight line between anchors and a chain that turns a corner meets itself exactly, instead of stepping out to a side lane and back on every vertical hop
- Cable lines are a little finer, and thin down further as you zoom out, so a wall no longer disappears under its own cabling at **Fit to View**

## Recent Changes In v0.43.0

**Signal and power cabling is drawn over the panels, and never over the text**

- Cable runs now sit **in front of the panel graphics** instead of behind them. On a normal wall every panel touches its neighbours, so the old behind-the-panels runs were completely hidden - all you ever saw was an arrowhead on each seam, and those landed straight on top of the panel labels
- The router keeps every run in the clear lanes a panel actually has: the strip **above the topmost label** for runs travelling left/right, and the **margin beside the centred text** for runs travelling up/down. Signal and power get a lane each, so the two never sit on top of one another either
- **Signal is a plain line with no arrows at all.** **Power is a plain line with one outline `>`** where it enters each panel - drawn open, never a solid arrowhead, and never repeated along a run
- Signal runs are drawn in a darker shade of their port colour. A signal run only ever crosses panels filled with its own port colour, so at full strength it was invisible now that it sits on top of them; the hue is unchanged, so a run still reads as that port's cable
- **The Panel Layout pages of the PDF use exactly the same routing, line weights and `>` marks as the workspace**, so the two finally match
- The port-number badges and the panel text are painted back over the cabling, so even where a run passes a badge the number stays readable

**Panel text is sized to the panel it is in**

- Panel labels were a fixed size, so they overflowed and overlapped once a panel got small - a zoomed-out workspace, or a narrow LED poster section, where the text was wider than the section itself
- They now step down to whatever fits the panel clear of the cable lanes, and are dropped entirely below the point where there is nothing readable left to fit. **A full 0.5m panel at 100% zoom or more, and every panel in the PDF, keeps exactly the text size it always had**

**Panel X / Y on hover**

- Hovering a panel in the Panel Layout shows **X** and **Y** for that panel, measured from the **top-left corner of the whole layout**, in mm and in content pixels. Like the row/column reference numbers, the figures come from the layout's true geometry, so Front View doesn't renumber anything

**Stray text gone from the PDF**

- The Panel Layout pages printed a `CX ... CY ...` line inside each panel whenever sub-screens or output-canvas positioning were in use - raw output-canvas coordinates that meant nothing on a cabling drawing. It is gone; the hover readout above is where a panel's position lives now

## Recent Changes In v0.42.0

**Wall Summary layout**

- **PowerPoint Content Setup is now its own panel** in Wall Summary instead of sitting inside Wall Details
- Dropped **LED wall aspect ratio** and **Native content resolution** from it - both are already shown in Wall Details, so the panel is just the slide size (plus the over-limit warning when it applies)
- **Area** moved to the bottom of Wall Details, under the resolution and ratio lines

**You can select and copy text again**

- Yes, this was something we had coded in, and it was a bug. The panel-canvas keyboard shortcuts are bound to the whole window, so **Ctrl+C copied the selected panels instead of the selected text** - the clipboard came back silently wrong wherever you tried to copy a figure out of the tool. **Delete** had the same problem: with text highlighted it deleted panels
- Ctrl+C, Ctrl+V and Delete now stand aside whenever text is highlighted, so copying works normally everywhere. With nothing highlighted they behave exactly as before, so copy/paste/delete of panels is unchanged. Undo, redo and the mode keys were never in conflict and are untouched

**LED Posters**

- Selecting the **LED Poster** panel type in the main tool now **greys out Rows and pins it to 1**, matching Quick Panel Layout - posters are always one high, and Apply Grid Size forces it too rather than trusting the box
- **+ Add Panel** becomes **+ Add Poster** for posters, and adds a complete poster (all eight sections) instead of a single loose section
- **Every poster is now its own sub-screen**, named `Poster 1`, `Poster 2`, ... with its own identity colour - whether it came from Quick Panel Layout, Apply Grid Size or + Add Poster. Each poster is a separate fixture with its own content feed, so it can be scoped, coloured, patched and given its own test pattern from the moment it is created. Names carry on past any sub-screens that already exist, so adding a second batch never reuses a name

## Recent Changes In v0.41.0

**LED Posters now split 2 wide x 4 high**

- Sending a poster to the main Layout Tool now splits it into a **2 x 4 block of eight 320 x 480mm / 172 x 258px sections** instead of four stacked full-width ones. The complete poster is unchanged - still 640 x 1920mm / 344 x 1032px - and Quick Panel Layout still works in whole posters, so wall size, resolution and power all come across exactly as before (2 posters read 1.28 x 1.92m / 688 x 1032px, 1,150 W / 5.00 A)
- All eight sections **stay grouped as one poster**: clicking any one of them selects the whole poster (the highlight now covers all eight, matching what move, rotate, copy, delete and sub-screen assignment already acted on), and never pulls in the poster next to it
- Stock Calculations still counts **complete posters**, not sections - 2 posters on the canvas is a required quantity of 2 against code 12199
- Per-section figures rescaled to eighths: **71.875 W / 0.3125 A** per section, still 575 W / 2.50 A per complete poster. Weight stays 0

**Signal and power port fixes found while making the above**

- **Panels per Signal Port was over the processor limit for posters.** It was set to 28 sections, which is 2,485,056 pixels - 382% of the 650,000-pixel-per-port ceiling the tool checks against. The figure had been worked out in posters and then written down as sections. It is now **8 sections = 1 complete poster = 355,008 px (54.6%)**, which is the real limit: two whole posters genuinely will not fit on one port
- **Panels per Power Outlet was capped at 21 for every panel type**, which is MG9's figure hard-coded in two places. It silently held posters at 21 sections (2.6 posters, splitting one poster across two outlets) and let MT be pushed to 21 panels - 22.9 A on a 16 A outlet. Each panel type now uses its own ceiling: MG9 21, MT 14, LED Poster 48 sections (6 whole posters, 3,450 W / 15.00 A)
- Changing Panel Type now **resets both port allowances to that panel's defaults** instead of carrying the previous panel's number over. The old behaviour kept any value that still "fit", so switching MG9 to MT left the port at MG9's 23 panels instead of MT's 39. Neither value is stored in a saved project, so nothing is lost
- Quick Panel Layout's distro sizing was **dividing whole posters by a per-section outlet figure**, so it sized circuits as though one outlet could take 48 complete posters. It now converts to whole posters and says "assumes 6 posters per outlet"

## Recent Changes In v0.40.1

- Filled in the **LED Poster** figures that were placeholders in v0.40.0:
  - **Power: 575.00 W per complete poster** (143.75 W per section, 0.625 A at 230 V - the same voltage basis the rest of the catalog uses). Four posters read 2,300 W / 10.00 A. Only one power figure was supplied, so average is set equal to peak: it can over-state a distro's load but never under-state it
  - **Weight stays 0** by instruction, so posters contribute nothing to the weight totals
  - Stock line added: **Tentec P1.86 LED Poster, code 12199** (10 in stock), counted in complete posters rather than the sections the grid holds
  - Outlet allowance set to 24 sections (6 whole posters = 3,450 W), derived from the 16 A x 230 V ceiling and rounded down to whole posters
- Fixed poster sections being counted as MG9 for rigging hardware: a catch-all branch gave every top-row poster section MG9's 1.9 kg fly bar and 1.5 kg sling, and put MG9 Hanging Bars in the stock list for them. Each panel type now carries its own hardware weight, so posters (0 kg) stay out of the rigging totals

## Recent Changes In v0.40.0

**LED Poster panel type**

- **LED Poster** is now a panel type in Quick Panel Layout: a complete poster is 640 x 1920mm / 344 x 1032px, and posters are locked to one high (the Rows control is disabled and says so)
- Sending posters to the main Layout Tool splits each one into **four stacked 640 x 480mm / 344 x 258px sections**, which is the unit the main tool patches and maps in. The wall size and resolution come across unchanged - 3 posters read 1.92 x 1.92m / 1032 x 1032px in both tools
- The four sections of a poster **stay grouped and behave as one physical poster**: clicking any section selects all four, and moving, rotating, copying, deleting or assigning to a sub-screen acts on the whole poster. Deleting one asks "Delete these 4 panels?" and takes the poster out in one go
- ⚠️ **The poster's weight and power draw are not in the catalog**, so they are currently **0** and contribute nothing to the Weight, Power and phase-load figures. They are deliberately not guessed - an invented figure would flow silently into rigging and electrical totals. Quick Panel Layout shows a warning while posters are selected. Send me the per-poster weight (kg) and power draw (max/avg W and A) and I'll fill them in
- There is also no stock line for posters yet: no matching equipment record was found in Rentman, so no code could be used. Send me the Rentman code and posters will appear in Stock Calculations like every other item

**PowerPoint Content Setup**

- The **Wall Details** panel now shows the recommended PowerPoint slide size, derived from the wall's own resolution - no inputs, no Calculate button, and it re-reads itself whenever the layout, panel type, orientation or resolution changes:
  - PowerPoint slide size, in cm to three decimals, sized so the longest edge is 100cm and the slide's proportions match the wall exactly
  - LED wall aspect ratio, reduced (2016 x 1176 shows as 12:7)
  - Native content resolution in pixels
- A 2016 x 1176 wall reads `100.000 x 58.333 cm`, `12:7`, `2016 x 1176 px`. On MT walls it uses the Recommended Content Resolution, since that is what content should actually be authored at. PowerPoint's 142.24cm ceiling is still checked and warned about, though the 100cm method stays well inside it for any realistic wall

**Layout**

- **LED Wall Setup** and **Wall Summary** can no longer be collapsed - they are the panels you work from and refer to constantly. Every other card still collapses as before

## Recent Changes In v0.39.1

- Swapped **Panels per Signal Port** and **Panels per Power Outlet** in the LED Wall Setup card so signal comes first, matching the order these are actually planned and patched in

## Recent Changes In v0.39.0

**Centre lines**

- The **Centre** label moved from above the wall to below it, in both the Panel Layout and the PDF. It was sitting directly on top of the metre ruler along the top edge and covering the measurements
- New **Show Sub-Screen Centre Lines** toggle in **Panel Layout -> Overlays & displays** (only offered when the project has sub-screens): a separate centre line for each sub-screen, each measured from that sub-screen's **own** panels rather than the whole wall, drawn and labelled in that sub-screen's colour. Independent of the whole-wall Centre Line toggle, and included in the PDF the same way

**Deployment**

- Choosing **Flown** as the Deployment Type now ticks every Additional Weight for you - Fly Bar, Sling & Shackle, Power cables and Signal cables. A flown wall carries all of them, and missing one silently under-states the rigging load. **Custom Weight is deliberately left unticked** (only you can supply that number), and every box stays freely un-tickable afterwards. Wired to the dropdown itself, so opening a saved project never overwrites weights you had switched off
- **Ground** deployment now adds **Temporary Fencing Weight** (code `12357`) to Stock Calculations at 3 per metre of wall width - a 3m wall asks for 9. Applies to MT ground-support walls too, not just MG9

**Moving Test Pattern**

- The output display is now asked for **every time**, listing every connected display rather than guessing. Previously a single secondary display was used silently, and a remembered choice could send the pattern to yesterday's monitor. Cancelling now cancels rather than opening it somewhere unasked-for
- Removed the **Change output display** button - with nothing remembered any more, there is nothing to change

## Recent Changes In v0.38.0

**Panel counts**

- One set of panel-count figures and one set of words for them, used identically in Quick Panel Layout, the Wall Summary, Stock Calculations and the PDF: **Required Panels**, **Spare Panels**, **Spare Panels - Rounded to Full Boxes**, **TOTAL Required Panels**
- **Spare Panels - Rounded to Full Boxes** takes required + spare up to a whole number of equipment boxes, because whole boxes are what leave the warehouse. The required panels already part-fill a box, so the spare only tops up whatever is left of it. On MG9 (boxes of 10): 45 required needs 4 spare, 45 + 4 = 49 rounds up to 50, so **5** spares go out for a total of 50 - a clean 5 boxes. Shaped Triangle/Curved panels are bought individually rather than boxed, so they aren't rounded at all
- **Spare Panels by Surface** is hidden completely when no LED surfaces are defined yet, instead of showing an empty table

**Project dates**

- The project **Start Date** and **End Date** now live in **LED Wall Setup**, directly under Project Name (they were in the Rentman card). Saved with the project as before, printed on the PDF's front page next to the project name, and used as the default window for Rentman availability checks

**PDF report**

- **Generate PDF** now opens a section picker first. Every section is ticked by default; untick anything you don't want. Sections that don't apply to the project (no sub-screens, no Rentman data pulled) aren't offered at all
- Fixed the **Signal & Power Ports In Use** page printing internal panel UUIDs in its Chain column (`1ebf57da-... -> 753e6d78-...`). It now prints the same panel reference used everywhere else in the tool - `R1 C6 -> R3 C1` - and the Power Outputs table gained the same Chain column. No UUID is exposed in the PDF anywhere

**Test patterns**

- Each sub-screen now generates and renders its **own independent moving test pattern**, confined to its own panels and running on its own phase - a wall with three sub-screens runs three separate patterns rather than showing one wall-wide animation sliced three ways. Each gets a boundary and name banner in its own colour. A wall with no sub-screens is unchanged
- **Moving Test Pattern** and **Download Moving Test Pattern** now ask which surface to show first. The choice is single-pick (radio buttons, **Full canvas** preselected), because the live view fills one display and a recording is one file. Picking a sub-screen renders that screen alone at its own resolution: a 2x3 sub-screen of a 6x3 wall comes out 336 x 504px, not 1008 x 504px with the rest black. A recorded sub-screen carries its name in the filename (`Untitled-Project-Left-Screen-front-test-pattern.webm`)
- **Test Pattern** now exports a package: the full-wall PNG plus a separate PNG per sub-screen (`Left-Screen-Test-Pattern.png`), each containing only that sub-screen's panels at that sub-screen's own true pixel resolution, with its colour on the border and name banner. A picker (everything ticked by default) chooses which ones to save

**Layout**

- Sub-screens now have a **colour** you pick (colour wheel or one of eight preset swatches), used for that sub-screen's outline in the Panel Layout and in its test patterns. Saved with the project
- The Centre Line setting moved from the top toolbar into **Panel Layout -> Overlays & displays**, and is now a single toggle that hides the centre line everywhere it is drawn - the layout on screen and the PDF's Panel Layout pages. There is no longer a separate "include it in the PDF" tickbox to contradict it

**Rentman**

- The separate **Rentman Integration** card is gone; **Get Current Stock from Rentman**, **Check Stock Availability by Date Range** and **Check Broken / Repair Equipment** now sit together at the top of **Stock Calculations**, next to the numbers they affect
- Availability defaults to the project's date range, and can be pointed at a different window from inside Stock Calculations without changing the project's own dates ("Back to project dates" restores it)
- New **Other Projects** column: how much of each item other Rentman projects need in the checked window. Click the number to expand the list - project number, name, status, dates and quantity - so how firm each booking is, is visible
- New **Broken / Repair** column: how many units are unavailable because they're broken or in for repair, read from Rentman's own repair records. Click to expand the serial number, repair status, date raised and Rentman's repair note. Counted by distinct serial, so two open faults logged against one panel is still one panel off the shelf
- Once Rentman has been checked, the Stock Calculations table reads: Equipment, Required, Spares, Spares Rounded to Full Box, Total Required, Rentman Stock, Other Projects, Broken / Repair, Available Stock, Result - where **Available Stock = Rentman Stock - Other Projects - Broken / Repair**, and the result is **OK**, **LOW** (covered by under 10%) or **SHORT n**. The same columns and a supporting **Other Projects & Repairs** detail page are in the PDF
- Confirmed stock quantities were already saved outside the project (browser-level, so they survive a reload and apply to every project) - unchanged
- `BOX-MG9` is not in the stock availability calculations - it was removed in v0.34.0 and stays out: once the spare quantity is rounded to complete boxes, a separate boxes line would double-count the same panels
- **This release needs the Rentman Worker redeployed** (`cd rentman-proxy && npm run deploy`) - Check Broken / Repair Equipment calls a new endpoint that older deployments don't have

**Housekeeping**

- `npm run lint` is clean. The app's entry point now picks the root *element* rather than assigning a capitalised component const, which is what react-refresh was flagging - the rule stays on everywhere, nothing is suppressed

## Recent Changes In v0.36.1

- Fixed the Moving Test Pattern's alignment circle rendering as a squashed ellipse on MT walls - it's now a true circle in the actual generated content-resolution image (a proper alignment reference should look genuinely round, not "round after an imagined future correction")
- Fixed panel column numbering (on-screen labels, the PDF, and the test pattern) double-counting MT panels - MT panels are 1000mm wide (not the 500mm module width MG9 uses), so 3 real MT panel-columns were previously labeled 5, 3, 1 instead of 3, 2, 1. Row numbering was already correct (MT's height does match the 500mm module) and is unaffected; the MG9-specific frame/floor deployment hardware quantities (which are genuinely module-based, not panel-based) are also unaffected

## Recent Changes In v0.36.0

- The Moving Test Pattern's status panel can now be hidden/shown three ways: click the panel itself to hide it, click anywhere else in the view to bring it back, or press `H` to toggle it either way
- Added a **Fit to Output** option (shown only when the test pattern doesn't fit the display 1:1) that deliberately scales the canvas to fill the display's width, aspect ratio preserved - clearly labeled "Scaled to fit output" rather than the usual not-1:1 warning, since this is an intentional choice, not an accidental mismatch. "Show Native Resolution (1:1)" switches back
- Fixed the bouncing MMS logo in the Moving Test Pattern rendering at the wrong aspect ratio on MT walls (it's live-preview decoration only, so it wasn't getting the same correction the rest of the pattern gets when eventually shown on real MT hardware) - it now displays correctly proportioned regardless of panel type

## Recent Changes In v0.35.0

- The live Moving Test Pattern is now rendered pixel-for-pixel at its true output resolution instead of being scaled to fit the browser window - a generated 1-pixel line is genuinely one pixel wide in the canvas. A new on-screen status panel reports LED Wall Resolution, Test Pattern Resolution, Display Resolution, Canvas Resolution, Device Pixel Ratio, Fullscreen state, and a measured (not assumed) **Browser -> Content Canvas: 1:1 / Scaled** verdict, re-checked live on resize, fullscreen changes, and device-pixel-ratio changes - with a clear `⚠ TEST PATTERN IS NOT BEING DISPLAYED 1:1` warning whenever it isn't. Press `H`, click the panel, or click anywhere else in the view to hide/show it. See [Pixel-Accurate Test Pattern](#pixel-accurate-test-pattern) below
- Opening the Moving Test Pattern now automatically opens and positions it on a second monitor (fullscreening it there) when one is connected and your browser supports it (Chrome/Edge over HTTPS or localhost, with permission) - the main app stays on your original screen. With 2+ secondary displays it asks which one to use and remembers your choice ("Change output display" clears that memory). Falls back cleanly to a plain window on unsupported browsers/contexts, with a one-time explanation of why
- MT's "Recommended Content Resolution" (double the vertical pixel count, to account for MT's non-square LED pitch) is now the resolution every generated test pattern actually uses - RGB/checkerboard, greyscale sweep, panel grid/outlines, alignment overlay and info text - not just a label as before. A dedicated `MT Vertical Content Mapping: 2:1` stat makes clear this is separate from the Browser -> Content Canvas pixel-accuracy check, so the two can't be confused for one another. Non-MT walls are completely unaffected (byte-identical output)

## Recent Changes In v0.34.0

- Removed the manual "map each item to Rentman equipment" step - this catalog's stock codes already match Rentman's own equipment codes, so **Get Current Stock** looks items up directly, no picker needed
- **Get Current Stock** now opens a review popup comparing Rentman's quantity against what's currently stored (old vs. new vs. difference, plus Rentman's own name for the item) - nothing is updated until you click Apply
- Replaced the single "Available (range)" column with a **Check Availability** popup: for each item, total Rentman stock, how much other projects need in the chosen date range, which projects (name, number, status, dates), and what's left over - makes a potential shortage easy to spot before committing to it
- Removed `BOX-MG9` and `BOX-MT` from Stock Calculations - by the time panel quantity is rounded to full boxes, that number already reflects complete boxes, so a separate boxes-required line added nothing
- While verifying this catalog's codes against the real Rentman account, found 4 that resolve to a different item than their local name suggests (a Dance Floor floor-reinforcement bar, floor taper pin, and tempered glass cover, plus the Triangle/Curved panel codes being swapped) - not corrected automatically since Rentman's item is the authority here, but the new comparison popup always shows Rentman's real name so a mismatch like this is visible before you apply anything

## Recent Changes In v0.33.0

- Added a **Rentman Integration** card to Stock Calculations: map each stock item to its matching Rentman equipment record (searchable picker), then click Refresh to pull live on-hand stock counts, overriding the built-in catalog numbers everywhere they're used (on-screen table, CSV export, PDF report, Shortfalls card)
- Set a date range in the same card to also show an **Available (range)** column - stock minus whatever's already booked on other Rentman projects overlapping those dates - in the on-screen table and PDF stock table
- Equipment mappings are saved in the browser (per-machine, not per-project); the date range is saved with the project file (v6 format - older saves still open fine)
- Requires a small separate one-time deployment (a Cloudflare Worker that holds the Rentman API token server-side, since this app is a static site with no backend of its own) - see [Rentman Integration](#rentman-integration) below. Without it, the card just shows "Not configured" and everything else works exactly as before
- Investigated pushing planned items to an existing Rentman project by number, as originally requested, but Rentman's public API has no way to add equipment to an existing project today (their own roadmap lists it as not yet shipped) - dropped from scope, may revisit if Rentman adds it

## Recent Changes In v0.32.1

- Reverted MT's spare ratio back to 0% (a deliberate catalog choice, not an oversight - v0.32.0 had mistakenly changed it to match MG9's 7%)
- Reworded the PDF's "Boxes: X (Y spare in boxes)" line to "Boxes: X (Y additional spare)" - the old wording used "spare" twice in a confusing way

## Recent Changes In v0.32.0

- Reworked how spare panels are shown: a new "Spare Panels by Surface" breakdown lists panels used, spare (7% of used), and spare rounded up to a full box, for each surface (each Sub-Screen plus "Unassigned" if any panels aren't in one, or just "Whole Layout" with no Sub-Screens) and each panel type (MG9 Standard, MG9 Corner, MG9 Triangle, MG9 Curved, MT) - with a subtotal per surface and a grand total when there's more than one. MG9 Standard/Corner round up to boxes of 10 and MT rounds up to boxes of 6; shaped panels (Triangle/Curved) are bought individually so their spare is added as-is, unrounded. MT panels now also get the same 7% spare allowance as MG9 (previously 0%). Shown in both the web app's Stock Calculations card and its own page in the PDF report

## Recent Changes In v0.31.0

- Fixed the Quick Panel Layout -> Main Layout Tool hand-off: when sending into a tab that already has an existing project (the Replace/Add prompt), the panel type now actually switches to match, and Columns/Rows now update too - previously only the fresh-project hand-off applied these correctly
- Cut the full PDF report's file size dramatically (a large wall could reach ~20MB) by rendering the embedded Panel Layout images at a fixed 300 DPI for their actual printed size on the page, instead of a flat pixel multiplier that scaled with the wall's real-world size - a huge wall always prints at the same page-sized image regardless of how big it is, so the old approach wasted enormous, invisible resolution on big projects. Typical/small projects are unaffected (the same ~300 DPI they already got); a 1200-panel test wall dropped from a projected ~20MB+ to well under 1MB with no visible quality loss. Also enabled PDF stream compression on both PDF exports

## Recent Changes In v0.30.1

- Relabelled the stock table's "Rounded" column to "Rounded + Spare" (on-screen and PDF) to make clear it's the order quantity, spare included
- Added Spare Panels Needed and Rounded To Full Boxes to the standalone Quick Panel Layout tool (on-screen and its PDF export), using the same per-panel-type spare ratio and box size as the main tool's own stock maths

## Recent Changes In v0.30.0

- Reworked the required-stock calculation so **Required** is always the raw quantity needed to build the wall (no spare folded in), and **Rounded** consistently adds each item's spare (plus packaging rounding, e.g. boxes of 10 for MG9 panels) - fixing several rows (MG9/MT/corner/shaped panels, power cable, signal cable) that previously showed an already-spared number in "Required". Stock shortfalls are now checked against the real order quantity (Rounded), not the bare required count. Items whose final order quantity comes out to 0 are hidden from the on-screen table, the PDF stock table, and the CSV export (which now exports the Rounded order quantity, not the raw required count) - the on-screen table also gained Spare/Rounded/Stock columns to match the PDF
- Fixed the PDF's "Signal Ports In Use" and "Power Outputs In Use" boxes silently dropping any ports past about 7 with no indication (easy to hit - VX2000 Pro alone offers up to 20 signal ports). Those boxes now show a compact in-use count, and a new dedicated "Signal & Power Ports In Use" PDF page lists every used signal port and power outlet in full, with no truncation

## Recent Changes In v0.29.0

- Correctly handle MT's non-square LED pixel pitch (3.9mm horizontal x 7.8mm vertical - every second LED row is physically missing on this transparent panel) everywhere a wall's resolution or aspect ratio is shown. For an MT-only wall, Quick Panel Layout and the main tool's Wall Summary now show **LED Wall Resolution** (the panel's real 256x64 pixel grid), **Recommended Content Resolution** (LED height doubled, so content authored at this resolution has the correct proportions), and a corrected **Physical Aspect Ratio** (from the wall's true physical size, not its raw pixel grid) - previously the "Aspect Ratio" stat was silently wrong for MT walls (e.g. showing 8:1 for a wall that's physically 4:1), and Quick Panel Layout's preview diagram and 16:9 content-area overlay were the wrong shape too. MG9 walls are unaffected, since its pixel and physical aspect ratios already match. The PDF exports from both tools, and the on-screen info text baked into the Test Pattern image/video, all reflect the same distinction for MT.

## Recent Changes In v0.28.2

- Added a "Show 16:9 content area" checkbox to the standalone Quick Panel Layout tool, controlling the dashed overlay box, its up/down nudge controls, its resolution stat, and its Full-HD warning, on both the web page and the PDF export - when shown, the PDF now also includes a short explanation of what the dashed box means
- Added an optional Project Name field to Quick Panel Layout, shown as a header on its PDF export and carried forward to become the main tool's Project Name when using **Send to Main Layout Tool**

## Recent Changes In v0.28.1

- Tidied the Quick Panel Layout PDF export: stats are now grouped under clear "Panel", "Power" and "Weight" section headers in a compact multi-column layout, replacing the old flat list that ran off the bottom of the page

## Recent Changes In v0.28.0

- Moved the orange power-port badge to sit directly beside the blue signal-port badge(s), all in the panel's top-left corner (was previously in the opposite corner) - neatly spaced, non-overlapping, in both the Panel Layout and the PDF Report
- Added an automatic weight estimate to the standalone Quick Panel Layout tool: panel weight, flying hardware (fly bars + slings/shackles, assuming a Flown deployment), estimated cable weight (power + signal, assuming cables snake left-to-right and alternate direction each row), and total flown weight - no manual rigging or cable input needed, clearly labelled as an indicative estimate rather than a certified rigging calculation, and included in its PDF export

## Recent Changes In v0.27.1

- Restored the shape-hugging signal/power chain-start ring indicators (removed in v0.27.0) - they're now shown together with the new port-number badges, not instead of them
- Moved each panel's info text (row/column label, assigned signal/power port, shape symbol) to the bottom of the panel, in both the Panel Layout and the PDF Report, to keep the top corners clear for the port-number badges
- The NovaStar processor model now defaults to VX2000 Pro for a new project instead of none selected

## Recent Changes In v0.27.0

- Added port-number badges to the first panel of every signal and power chain, in both the Panel Layout and the PDF Report: a blue circle (top-left) with the signal port number, and an orange circle (top-right) with the power port number - replacing the old shape-hugging "chain start" ring outlines. Badges are drawn outside the panel's own rotate transform so the digit stays upright and legible even on a rotated panel, and follow the panel when it's moved, rotated, or the layout is exported
- Added backup signal port numbering: with **Do backup signal loop** enabled, the chain's last panel also gets a badge with the backup port number - the second half of the available signal ports backs up the first half (e.g. port 11 backs up port 1 on a 20-port setup; port 6 backs up port 1 on a 10-port setup)
- The number of selectable signal ports now follows the selected NovaStar processor model (10 for VX1000 Pro, 20 for VX2000 Pro) instead of always offering 20 regardless of processor; falls back to 20 when no processor is selected
- With the backup signal loop enabled, the second half of the port range is reserved for backups: hatched and unselectable in the Signal Patching panel, and automatically excluded from Auto Snake / manual port assignment, so primary signal-port capacity updates automatically

## Recent Changes In v0.26.0

- Added a **Power** summary to the standalone Quick Panel Layout tool: total power draw (max/avg W and A) plus, for both a 32A and a 63A distro, the circuits needed, distro units needed, and percentage of safe capacity used - reusing the same per-panel power spec and safe-panels-per-outlet defaults as the main Layout Tool, also included in its PDF export
- Added +/- stepper buttons next to the Width (m) and Height (m) fields in Quick Panel Layout, matching the existing Columns/Rows steppers, for easier use on mobile/touch

## Recent Changes In v0.25.1

- Fixed a real regression from v0.25.0's DPI-aware live Moving Test Pattern rendering: the RGB checkerboard tiles no longer lined up with panel boundaries (each flat-coloured square could span multiple panels). Root cause was `drawTestPatternFrame` resetting the canvas transform to a hard-coded identity to draw its pre-rendered pattern layer, which only stays correct when the canvas is 1:1 with its content (true for the recorded video/PNG/PDF exports) - once the live view started rendering at a devicePixelRatio/fit-to-window scale, that hard reset drew the pattern at the wrong size. It now resets to whichever transform the caller had active instead of a literal identity
- Regrouped the Panel Layout toolbar: **Patch**, **Select**, **Move** and **Clear Patching** now live together in one "Panel tools" group (previously split across two groups), with the Snap/Move joined group/Allow overlaps options staying alongside Move

## Recent Changes In v0.25.0

- Doubled the bouncing MMS logo's size (now ~1/2 a standard panel's native pixel width, up from 1/4) and halved its movement speed in the browser-only live Moving Test Pattern
- The "Include Centre Line" PDF export option now defaults to enabled and sits at the end of its toolbar row
- Replaced the separate WebM/MP4 download buttons with a single "Download Moving Test Pattern" button that opens a format-choice dialog (WebM or MP4, with a clear warning that MP4 takes significantly longer to encode) with Download/Cancel
- Fixed blurry/sub-pixel text, lines and arrows in the live Moving Test Pattern tab: the canvas's backing store is now sized to the actual physical device pixels it's displayed at (CSS size x devicePixelRatio), not just the wall's native resolution, so high-DPI displays and non-1:1 window scaling no longer blur or alias thin strokes
- Fixed rotated-panel snapping: panels rotated to a custom angle (not just 0/90/180/270) now snap and join along their own true rotated edges instead of silently snapping as if they were unrotated. This was a real bug in `panelWorldAnchors` (it rounded rotation to the nearest 90deg before computing connector anchor positions) affecting individual panels, multi-selected groups, custom angles and imported rotated panels alike; covered by a new regression test
- Added a **Fit to View** button that zooms the Panel Layout workspace so the entire project - including a wide or tall imported layout - is visible at once, with the scroll position reset to the origin; confirmed the existing scrollable workspace already expands and scrolls correctly for any project size
- Reorganised the Panel Layout toolbar into labelled groups (Selection & editing, Move/align & snap, Rotation & transforms, View/zoom & navigation, Overlays & display) instead of one long unlabelled row - all existing controls preserved, including moving the Front/Back View and Centre Line toggles down from the card header into their matching groups

## Recent Changes In v0.24.0

- Rounded the measurements shown in the Panel Layout header (and the Wall Summary/PDF size lines) to at most 2 decimal places with trailing zeros trimmed, instead of the raw unrounded floating-point value
- Expanded panel rotation: dedicated 45° and 90° buttons plus a custom-angle input, on top of the existing keyboard shortcut. Rotating a multi-selection spins every selected panel in place by the same amount, so their arrangement and spacing relative to each other never changes
- Fixed layout imports from the Creative Layout Tool silently snapping every panel's rotation to the nearest 90° and mis-mirroring square panels' rotation under the front-view flip - both bugs only showed up once panels could be rotated to non-cardinal angles. Imported panels now keep their exact source rotation and orientation
- The NovaStar Processor Configuration section is now collapsed by default and moved to the very bottom of the sidebar, out of the way during normal layout work
- Added copy/paste for selected panels: Ctrl/Cmd+C copies, Ctrl/Cmd+V enters a paste-placement mode with a dashed preview that follows the cursor and snaps to the grid/nearby panels, click to place, Escape or right-click to cancel. Pasted panels are auto-selected and keep their source spacing, rotation, panel type and arrangement (patching is left unassigned, same as importing a layout)
- Applying a new grid size over an existing layout now asks for confirmation ("Remove Panels and Apply Grid" / "Cancel") instead of silently wiping every panel
- The Moving Test Pattern's per-panel row/column direction indicators are now drawn as vector arrows (explicit stroke width with a floor, arrowhead scaled to match) instead of relying on a font glyph's own internal strokes, which could shrink below a visible pixel width at small panel sizes or when the output is scaled
- Added a small bouncing MMS logo (DVD-screensaver style) to the browser-only live Moving Test Pattern view - about a quarter of a panel's native width, aspect-ratio preserved, stays fully inside the canvas, never appears in the recorded video or PNG/PDF exports
- Added a toggleable vertical centre-line indicator to the Panel Layout workspace, with a "Centre" label and an "Include Centre Line" option for the PDF export. The centre is computed from every active panel's true rotated outer bounds, not just the axis-aligned wall bounding box, so a panel spun to a non-cardinal angle is still accounted for correctly
- Selecting a sub-screen now fully hides every panel not assigned to it (previously they stayed visible, dimmed and locked) - the active sub-screen's name is shown clearly above the workspace, and an "All Screens" option (renamed from "Canvas View") returns to the complete layout instantly. No panel data is ever changed by switching the visible screen
- Renamed **Video Test Pattern** to **Moving Test Pattern** and dropped "PNG" from **PNG Test Pattern** (now just **Test Pattern**) everywhere the labels appear - buttons, tab titles, help text and the README
- The browser's live Moving Test Pattern tab now anchors the LED canvas to the top-left corner and scales it to fit the window on both axes (never stretched, cropped, centred, or auto-rotated between portrait/landscape), recalculating on resize; unused space fills with a plain black background instead of centring the canvas
- Added a **Download MP4** option next to the existing WebM download. MediaRecorder can't produce MP4 directly in most browsers, so this records the same WebM as today and then transcodes it to H.264 MP4 in the browser via a lazily-loaded ffmpeg.wasm (only fetched the first time this button is used - roughly 30MB, entirely separate from the app's normal bundle)

## Recent Changes In v0.23.0

- Added **NovaStar Processor Configuration File Export**: pick `VX1000 Pro` or `VX2000 Pro` in LED Wall Setup (with a live capacity-vs-usage readout beside the selector), assign a video input per sub-screen or one input for the whole canvas via a mode toggle in Output Canvas, then generate a real `.uprj` file from a new "NovaStar Processor Configuration" section showing a full pre-download summary (processor, surface, canvas/screen resolution, panel/sub-screen counts, Ethernet outputs used, pixel load per output, input assignments) plus any validation errors (which block download) or warnings. The file format - a custom envelope wrapping an embedded SQLite project database, including its checksum algorithm - was reverse engineered byte-for-byte against real NovaStar-exported project files rather than guessed; cabinet/panel positions and signal-port patching order are written into the LED screen's own native pixel grid exactly as NovaStar itself represents it, verified byte-for-byte against real reference exports covering gapped/irregular walls, multiple Ethernet outputs, and multi-sub-screen layouts. Output-canvas/sub-screen placement is not yet reflected in the exported cabinet positions (only in the video-routing side) - a known gap for a future pass. Settings JSON bumps to v5 to carry the new processor/input fields; older saves load with no processor selected, same as always
- Output Canvas and Stock Calculations are now collapsed by default
- Fixed the PNG and animated/video test patterns not locking panels to their true positions on walls with a missing/removed panel in the middle of a row: panels were packed tightly left-to-right in array order (silently closing up any gap), shifting every panel after the gap out of position. Both now place each panel at its real relative pixel offset, so a physical gap shows as empty space instead of squeezing later panels together - uniform, gap-free walls render identically to before. The PNG export also had its own separate copy of this same (still-buggy) positioning logic despite the live/video test pattern already having been fixed for it previously; it now shares the one, corrected implementation
- Fixed the test pattern's per-panel row/column corner labels reading from the panel's raw back-view position instead of the front (mirrored) view the pattern always renders in - column numbers were backwards versus what's on screen. The rendered top-left panel now always reads row 1, column 1, and the centred info-text block is smaller
- The Panel Layout workspace's (and PDF's) height ruler now reads bottom-up - 0m at the wall's base, increasing upward - matching how a physical wall is measured and built; tick positions are unchanged, only the printed labels. Added compact direction arrows to the Columns/Rows field labels
- Added an automated test suite (Vitest, `npm test`) covering the NovaStar export pipeline - envelope/checksum round-trips, both processor models, irregular and multi-sub-screen walls, patching-order preservation, and golden-file comparisons against real NovaStar-exported project files

## Recent Changes In v0.22.2

- Fixed the PNG and animated/video test patterns rendering MT panels at the wrong resolution: both always sized the whole canvas and every panel using MG9's mm-to-pixel ratio (168px per 500mm), which happened to be correct for MG9 but gave MT panels 336x168px instead of their true native 256x64px. Both exports now place each panel using its own native `pixW`/`pixH` from its panel type, accumulated per row band - the same algorithm already used for the "Resolution" stat in Wall Summary - so the exported canvas size and every panel's pixel footprint always match what the app reports on screen. MG9-only walls are unaffected

## Recent Changes In v0.22.1

- Fixed the animated/video test pattern's RGB checkerboard: the colour tile was always sized to an MG9 panel's pixel width, so on MT walls each panel (physically twice as wide) showed two colours split down its middle instead of one solid colour. The tile now matches the actual panel footprint on the wall (derived from the most common panel type present), so MT panels correctly show one colour across their full width and MG9 walls are unaffected
- Reworked the test pattern's text legibility and sizing: the centre info block (project name/stats) now has a black stroke behind its white fill so it stays readable where the alignment cross-hatch overlay crosses it, is much larger, and the per-panel row/column corner labels are now half their previous size. The alignment cross-hatch overlay itself is now more transparent

## Recent Changes In v0.22.0

- Added **Sub-Screens**: create/rename/delete named panel groupings, assign/reassign/remove panels, select-all-in-screen, and a Canvas View showing the whole layout at once. Selecting a sub-screen dims and locks every panel outside it and scopes all calculations (panel count, resolution, weight, power, stock, port usage), manual/auto patching, and PDF/PNG/video exports to just that screen. A labelled boundary outline is drawn in the workspace for each sub-screen. Projects with no sub-screens behave exactly as before, and old save files load unchanged
- Added **Output Canvas Positioning**: configurable output resolution (common presets plus custom width/height), per-sub-screen (or whole-layout) placement via numeric X/Y entry, dragging, arrow-key nudging, or align/centre/snap-to-edge tools, a scaled live preview, and warnings for out-of-bounds or overlapping screens. Each panel's final canvas-space X/Y is computed from its sub-screen's position plus its own placement, independently of the physical millimetre layout
- Auto-patching, manual patching, and the clear/match-power actions now operate only on the active sub-screen's panels (shared signal/power port pools are respected, and other sub-screens' patching is never touched or reset)
- PDF export gains a per-sub-screen summary page (name, resolution, physical size, canvas position, panel count, ports in use) whenever sub-screens exist
- Fixed sub-screen creation jumping straight into the new (empty) sub-screen, which visually collapsed the workspace and dimmed every existing panel until something was assigned to it - creating a sub-screen now leaves the current view untouched
- Reorganised the layout: removed the legacy "LED Surface / Sub-Screen Name" field, moved the Sub-Screens panel below Wall Summary, moved Output Canvas below Power Outputs (and made it always visible, no longer behind a toggle), and moved panel-to-sub-screen assignment controls next to Undo/Redo with a live selection count
- Every major section of the UI is now collapsible via its header

## Recent Changes In v0.21.2

- Quick Panel Layout: added an `Export PDF` button - a one-page landscape summary with the panel type/grid/wall size/resolution/aspect ratio/16:9 content-area stats plus a to-scale diagram of the wall (grid lines, metre rulers and the 16:9 overlay), matching the on-screen preview

## Recent Changes In v0.21.1

- Quick Panel Layout: added a `Clear` button (resets to 1×1 MG9), metre rulers along the top and left of the preview, and up/down buttons to nudge the centred 16:9 content-area box vertically within any available slack
- Quick Panel Layout: added Width (m) / Height (m) inputs alongside the Columns/Rows counters, kept in sync both ways - width steps in 0.5m increments for MG9 / 1m for MT, height always steps in 0.5m (both panel types are 0.5m tall)
- Quick Panel Layout: reworded the preview caption from "Dashed box = centred 16:9 content area" to "Dashed box = 16:9"

## Recent Changes In v0.21.0

- Added **Quick Panel Layout**: a standalone panel-count calculator, opened in its own browser tab (`Quick Panel Layout` toolbar button), independent of any open project. Pick MG9/MT and columns/rows and see live wall size, pixel resolution, panel count, aspect ratio, a centred 16:9 content-area overlay, and a warning if the wall (or the 16:9 area) is below 1920×1080. A `Send to Main Layout Tool` button hands the chosen grid off and jumps to the main planner, which applies it on load (replacing the canvas if it already has panels, after a Replace/Add to canvas/Cancel prompt)
- The main LED Cabling Planner now starts with an **empty canvas** instead of an auto-generated 24×8 grid - build a layout via `Apply Grid Size`, import, open a saved project, or send one in from Quick Panel Layout

## Recent Changes In v0.20.2

- Fixed the on-screen signal/power chain-start ring indicator rendering as a square on MT (or any non-square) panels instead of following the panel's true rectangle. The ring's SVG had no explicit width/height, so browsers fell back to its viewBox's 1:1 intrinsic aspect ratio when sizing it as an absolutely-positioned element, overriding the intended stretch-to-fill. Now explicit, always matches the panel's actual shape

## Recent Changes In v0.20.1

- Removed the signal/power chain-start ring indicators from the PNG test pattern - it's a clean per-panel pixel map now, not a patching diagram. PDF and on-screen views still show them

## Recent Changes In v0.20.0

Consolidates and confirms the PNG shape-rendering fix: verified consistent between the PNG test pattern and the PDF's front view across triangle panels at all four rotations (0/90/180/270) and the corner panel's hatch pattern, in addition to the curved-panel fix already in v0.19.2.

## Recent Changes In v0.19.2

- Fixed a bug where curved (MG13) panels rendered with the wrong corner cut in the PNG test pattern - the PNG traced a curve path geometrically opposite to the one used on-screen and in the PDF, making rotated curved panels look incorrectly oriented. The PNG now always shares the exact same shape-tracing logic as the PDF/screen, so per-panel rotation matches the PDF's front view exactly
- Removed the rotate icon (🔄) and all per-panel signal/power port info from the PNG test pattern; panels now show only their row/column label and shape symbol

## Recent Changes In v0.19.1

- The outer-extremity outline is now a single, thicker white line (3px) instead of two separate 1px lines with a gap between them

## Recent Changes In v0.19.0

- Panel alignment outlines are now pixel-snapped and drawn as a crisp true 1px line (rect/MT/corner panels get an exact strokeRect fast path; shaped panels keep their straight legs crisp)
- Removed the black outline around the wall info text; removed all per-panel signal/power port labels from the test pattern
- Added a "LED Surface / Sub-Screen Name" field (alongside Project Name), saved with the project; both names are shown centred on the wall when defined, with no placeholder when empty
- Panel location labels moved to the top-left corner of each panel as two lines (`↓row` / `→col`), consistently positioned regardless of shape or rotation
- The full-screen live view now defaults to true 1:1 pixel mapping (centred if smaller than the window, scrollable if larger) instead of stretching to fit; any keypress toggles an optional scaled-to-fit preview
- Added a double 1px white outline around the true outer extremity of the whole assembled LED surface (not per-panel), accurately following triangular/curved/irregular outlines and ignoring internal panel-to-panel seams

## Recent Changes In v0.18.0

Animated test pattern tweaks:

- Split into two dedicated buttons: **Video Test Pattern** opens a pure full-screen canvas in a new tab - no header, no buttons, no text outside the LED canvas itself; **Download Video Test Pattern** records and downloads the WebM directly from the main app, no tab required
- The moving greyscale gradient is now a single large sweep spanning the whole wall corner-to-corner, instead of several smaller repeating bands
- Removed the info panel's background box and the "Test: ..." description line; the remaining wall info (resolution, physical size, panel count, grid) is now centred on the wall
- Added a corner-to-corner alignment cross and a centre circle (diameter equal to the wall's height) as a geometry reference for spotting warped, offset or stretched panels
- Fixed washed-out/blocky WebM exports by giving the recorder a much higher, resolution-scaled video bitrate instead of the codec's low default

## Recent Changes In v0.17.0

Added an animated LED wall test pattern for spotting orientation, patching and alignment errors that a static swatch can't reveal:

- New "Animated Test Pattern" button opens a live, looping canvas view in its own tab, rendered at the wall's exact configured pixel resolution
- RGB checkerboard: every panel shows a solid red, green or blue test colour in a diagonally staggered arrangement (never a blended rainbow), sliding smoothly left-to-right and cycling red -> green -> blue -> red
- A moving diagonal greyscale brightness sweep plays across the whole wall at the same time, continuous across every panel boundary (not restarting inside each panel), without introducing colour or making panels hard to identify
- 1px white outlines follow each panel's true shape (rectangle/triangle/curve) and rotation
- Every panel is labelled (row/column, signal port, power port) in white, correctly positioned even on rotated or shaped panels
- A small on-canvas info panel shows resolution, physical size, panel count, grid size and the active test description
- The whole animation loops seamlessly every 20 seconds (verified bit-for-bit identical at the loop boundary) and always renders Front View, matching the PNG test pattern's convention
- "Download Video (WebM)" records exactly one loop as a native WebM file (no extra dependencies - browser MediaRecorder/canvas.captureStream) that plays back looped with no visible seam
- Works for uniform grids and freely placed/imported non-uniform layouts, including mixed MG9/MT and rotated/shaped panels, by defining the animation in wall pixel-coordinate space and revealing it through each panel's own clip mask

## Recent Changes In v0.16.0

Non-uniform layout overhaul (Stages 2-4) plus a round of fixes and new editing features, delivered as staged local commits:

**Free panel placement + import**
- Panels are no longer a fixed grid: place, drag, rotate, snap, join and multi-select panels freely, with overlap warnings and a live snap/join guide
- Import projects from the Creative Layout Tool, with a preview (name, panel mix, wall size) before replacing or adding as a new project
- Imported projects are interpreted and displayed as **Front View**, matching the original Creative Layout Tool design exactly (position, shape and rotation), instead of the app's default back/wiring view
- New save format v2 (free mm-positioned panel list); legacy grid-format settings files still open and migrate automatically
- Automatic letter-shaped patching (bottom-up, fork-aware) for text/logo-shaped layouts

**Editing and safety**
- Deleting panels now prompts with **Remove Panel**, **Mark as Inactive**, or **Cancel** — inactive panels stay visible (dashed) in place but are excluded from totals, patching and exports
- Keyboard shortcuts `S` (Select), `M` (Move), `P` (Patch) documented in Help, alongside the existing shortcuts

**Signal/power cable rendering**
- Cable lines now draw behind panels with a thin black outline; arrowheads draw in front, also black-outlined, and always point in the true signal/power direction (including when adjacent panels touch edge-to-edge)
- A selected panel is brought to the front, above cable lines, so its info stays readable
- Orthogonal (90°) cable routing everywhere: on-screen, PDF and PNG test pattern
- Snap/join logic ported from the Creative Layout Tool (connector-anchor based, shape/rotation aware)
- Signal/power chain-start indicator outlines now follow the true panel shape (triangle/curve/rect) at any rotation

**PNG test pattern export**
- Always renders Front View regardless of the on-screen toggle, matching what an observer sees standing in front of the finished wall
- No longer includes cable-routing lines or arrowheads
- Fixed panel alignment (true mm positions, no band-packing offset) and rotation accuracy
- Excludes inactive panels

**UI**
- New design-system `Button` component with clear active/selected states across all toolbar controls
- Panel Type control moved above Apply Grid Size; added Clear All Panels
- Renamed the import button to "Import Project from Creative Layout Tool"
- Dashed, wall-aligned background grid (1m major / 0.5m minor lines)

## Recent Changes In v0.15.0

Stage 1 of the non-uniform overhaul: interface refresh (data model unchanged).

- New shared design-system `Button` with consistent intents (primary / secondary / ghost / danger / success) and a clear active/selected state (bright fill + ring), replacing ad-hoc per-button colours
- Tools and modes now show an unmistakable active highlight: Signal / Power patch mode, Select mode, and view flip
- Controls grouped into labelled sections (Patch mode, Auto patching, Documentation & exports, Import & save, Selection & editing) with status chips for the active mode
- Cleaner cards, spacing and typography; extracted UI primitives into `src/components/ui.tsx`
- Cleanup: fixed the long-standing `useState` type warnings (typecheck now clean), tightened `patchMode`/`snakeDirection` types, removed a shadowed variable

## Recent Changes In v0.14.0

- Added chain-start indicators drawn alongside (never replacing) the existing panel outlines
- Blue outline on the first panel of each signal chain; orange outline on the first panel of each power chain
- A panel that starts both chains shows both outlines as clearly separated concentric rings (blue outer, orange inner)
- With "Do backup signal loop" enabled, the blue outline is also added to the last panel of each signal chain to show where the backup loop connects
- Indicators appear everywhere the layout is drawn: the live editor, printed page, PDF export, and the PNG test pattern

## Recent Changes In v0.13.0

- You can now mix `MG9` and `MT` panels in one wall. The layout is a 0.5m module grid: MG9 fills one module, MT spans two side-by-side modules
- Select Mode has a panel-type dropdown (MG9 / MT) to convert the selected panels; MT takes the module to its right and is blocked at the right edge
- Live stats are per panel type: panel count, weight, power/amps, and pixel resolution all use each panel's own profile
- Signal/power patching, auto-snake, and Match Power To Signal Pattern all treat an MT as a single panel and route cabling to its true edges
- Stock table is split by type and combined: MG9 (panels + variants) and MT each use their own catalog, spare ratio, box size, and hanging bar; MG9-only hardware (reinforcement, corner connectors, ground/floor frames) counts MG9 panels only
- PDF layout draws MT as a wide `(MT)`-labelled panel, with a mixed panel-type summary; PNG test pattern places each panel at its native pixel size (MG9 168x168, MT 256x64)
- Opening an older all-MT settings file migrates it onto the new module grid automatically

## Recent Changes In v0.12.0

- Added a `Match Power To Signal Pattern` button next to `Power Patch Mode`
- Power patching can now follow the existing signal patch: panels are powered in signal order (signal port, then sequence)
- Power plugs line up with the signal ports - each signal port starts on a fresh plug, giving a clean 1:1 plug-to-port mapping when a port fits in one plug
- Large signal ports spill onto consecutive plugs in order, and the tool still respects the power panel-count and 16A-per-plug limits

## Recent Changes In v0.11.0

- `MT` panels now render to their true 1m x 0.5m shape (2:1 wide rectangles) in the live panel layout instead of as squares
- PDF layout pages now draw `MT` panels as the same wide rectangles, with patching and power arrows routed correctly between them
- Cell width now scales from each panel profile's real-world width/height, so `MG9` stays square and `MT` is twice as wide as tall
- Confirmed the PNG test pattern exports at the correct native pixel ratio (256 x 64 per `MT` panel)

## Recent Changes In v0.10.1

- Restored patching arrows in the live panel layout and PDF layout pages
- Switched the test-pattern export from JPG to lossless PNG at true wall pixel dimensions
- Kept patching arrows and first-power markers out of the PNG test pattern only
- Improved Select Mode so panel editing does not accidentally patch panels
- Added undo/redo controls and shortcuts for layout edits
- Added a Help button with shortcut and workflow guidance
- Linked the version badge to this changelog
- Improved MG12 triangle, MG13 curved, and MG9 corner-panel drawing
- Changed corner-panel text to `Corner` and reduced corner hatching in the PNG export
- Updated connector stock values for `12260` and `12258`
- Reworked the PDF first page to show the full stock summary table before adding overflow stock pages

## Recent Changes In v0.10.0

- Added MG9-compatible special panel variants: MG12 Triangle, MG13 1/4 Curved, and MG9 LED Corner Panel
- Added persisted per-panel variant and rotation data in settings files
- Added drag selection for editing multiple panels at once
- Added multi-panel actions for changing panel type, rotating, clearing patching, deleting, and restoring panels
- Added keyboard shortcuts: `Delete` removes selected panels, `R` rotates, `C` clears selected patching, and `Escape` clears selection
- Removed the visible `Removed` label from deleted panels so holes stay visually blank
- Added a selected-port clear action for clearing the active signal port or power plug
- Backup signal loop now doubles the effective processor signal-port count while keeping the visible primary patch path readable
- Added corner-panel stock logic for corner panels, flat connectors, and corner connectors
- PDF Stock Summary now includes item names, item codes, required quantities, spare stock, rounded quantities, stock, and net stock
- Added JPG test-pattern export using the front view, true wall pixel dimensions, existing panel labels, port colors, panel shapes, hatches, and a 1px white border

## Recent Changes In v0.9.0

- Back view is now the default panel-layout view
- Panel layout clearly shows the current and alternate view
- PDF export now includes both `Back View` and `Front View` layout pages
- Load bars now stay orange when near the limit and only turn red when overloaded
- Added `LED Wall Deployment Settings`
- Added `Do backup signal loop`, enabled by default
- Added deployment types: `Flown`, `Ground`, `No Support`, `Floor`
- Backup signal loop now doubles `15m Signal Cable`
- Backup signal loop now adds `SEETRONIC SE8FF-05 F/M - F/M Joiner` per signal port with fallback to `SEETRONIC F/M - F/M Cable`
- Added MG9 ground and floor deployment stock calculations
- Settings export/import now includes deployment type and backup signal loop
- Added `Loop together` auto-snake preset
- Updated signal-port colors to match the NovaStar Unico look more closely
- Added stock CSV export and simplified stock table columns
- Added `12317 LED Prod Case` to every project stock list
- Added orange power-run start outlines on panel layout views
- Removed the dark background from panel-layout PDF exports
- Added removable and restorable panels for non-grid wall shapes
- Removed panels now skip patching, counts, stock, power, and support math
- Added per-panel `Clear Power And Signal Patching`, `Delete Panel`, and `Restore Panel` actions

## Local Development

```bash
npm install
npm run dev
```

Or double-click:

```text
start-local.bat
```

## Testing

```bash
npm test
```

Runs the Vitest suite (`src/**/*.test.ts`), including the NovaStar export tests, which compare generated files byte-for-byte against real NovaStar-exported reference projects checked into `src/novastar/__fixtures__/`.

## Production Build

```bash
npm run build
npm run preview
```

## GitHub Pages Deployment

This repo includes [`.github/workflows/deploy-pages.yml`](./.github/workflows/deploy-pages.yml).

To publish:

1. Push this folder to the GitHub repository.
2. In GitHub, open `Settings` -> `Pages`.
3. Set the source to `GitHub Actions`.
4. Push to `main` or rerun the Pages workflow.
5. Wait for the `Deploy GitHub Pages` workflow to finish.

The site uses a relative Vite base path so it works on repository Pages URLs.

## Test Pattern Generator

Live at `https://underdog1234.github.io/LED-Cabling-Web-App/generator.html`
(the **Test Pattern Generator** button in the planner opens it). It needs no
project and no panels: start from a resolution.

**Canvas and sub-screens**

- One main canvas of any whole-pixel size (presets for 1280 x 720, 1920 x 1080,
  1920 x 1200, 2560 x 1440, 3840 x 2160, 4096 x 2160, their portrait versions and
  the planner's wide canvases, or anything you type, with an optional aspect
  lock and your own named presets). Unused canvas is black, or any colour
- Any number of **sub-screens**, each with its own name, colour, resolution,
  X/Y, test pattern, pattern settings, overlays and animation offset - so
  different patterns run side by side. Add, duplicate, delete, rename, hide and
  lock them; overlap them and set the layer order
- Drag to move, drag the handles to resize (Shift keeps the ratio), or type
  exact numbers. Arrow keys nudge the selection (Shift for the bigger step;
  both steps are settable). Every position and size is a whole pixel
- Snap to canvas edges, canvas centre, other screens and an optional grid, with
  guides shown while you drag (hold Alt to place freely)
- Select several screens (Ctrl/Cmd-click, Shift-click, or drag a box) and
  align left / centre / right / top / middle / bottom to the canvas, the
  selection or a reference screen; match a reference's size; distribute evenly
  (gaps differ by at most one pixel when the space doesn't divide); space them
  with a fixed gap; or arrange in a row, a column or a grid with pixel gaps
- The planner's Output Canvas warnings: off-canvas, negative positions,
  overlaps, and a layout bigger than the canvas
- Optional LED maths: a calculator for resolution from panel rows, columns and
  panel pixel size, and a per-screen panel grid that drives the LED Moving
  Pattern. NovaStar processor input assignment (whole canvas or per screen) is
  kept too
- **Import LED Planner project** opens a saved planner file: its Output Canvas
  resolution, each sub-screen at its canvas position and footprint resolution,
  running the planner's own Moving Test Pattern, with its processor inputs

**Patterns**: Screen ID & Resolution, SMPTE and EBU bars, colour checker
patches, grid / crosshatch, circles with safe areas and aspect frames,
checkerboard, greyscale steps, gradient ramps, solid fields, colour field
cycle, moving bar, timecode with sync flash, pixel structure, zone plate,
fictional faces (illustrations from a seed, across twelve skin tones, each
labelled fictional), the LED Moving Pattern, and the planner's own layout
pattern for imported screens. Animated patterns all loop exactly on the
canvas's loop length.

**Output and playback**

- Play / pause / restart, a playlist of steps (each step remembers every
  screen's pattern and settings, for a set number of seconds, looping or not)
- **Output window** opens the whole canvas or one screen, at native
  resolution, on the display you choose (same display picker as the planner);
  it follows the editor live. **Fullscreen** shows it full screen in place.
  Both report whether the output is really 1:1, with Fit to output and the
  live-only bouncing logo

**Downloads** - choose before exporting:

- **Entire canvas** at its exact resolution with every visible screen
- **One sub-screen** at its own resolution - rendered on its own, so a
  neighbour that overlaps it never appears in the file
- **Selected sub-screens** as separate files in one ZIP

as a PNG still (at a chosen point in the loop), a WebM loop, or an MP4 loop
written with the planner's delivery settings. Files are named
`<screen>_<pattern>_<width>x<height>.<ext>`.

**Saving**: the working setup is autosaved in the browser; named
configurations can be saved and reloaded there, and any setup saved to or
opened from a `.testpattern.json` file - every screen's position, resolution,
pattern, settings, overlays and layer order, the playlist and your presets.

## Pixel-Accurate Test Pattern

The live Moving Test Pattern's canvas is always sized to exactly its Test Pattern Resolution (the LED wall's native resolution, or MT's doubled Recommended Content Resolution) - never scaled or shrunk to fit whatever window/display it's shown in by default. If that resolution doesn't fit the current display, the pattern overflows/clips rather than being resized - a mismatch is always reported, never silently resolved by scaling.

The status panel in the corner reports what's actually happening, measured from the real DOM rather than assumed from the numbers set up - click the panel to hide it, click anywhere else in the view to bring it back, or press `H` to toggle it either way:

- **Physical LED Resolution** / **Recommended Content Resolution** (MT only) / **Test Pattern Resolution** - what's being generated
- **Display Resolution** / **Canvas Resolution** / **Device Pixel Ratio** / **Fullscreen** - what the browser/display are actually doing
- **Browser -> Content Canvas: 1:1 / Scaled** - is the canvas genuinely landing 1 pixel per physical display pixel right now? Re-checked on every resize, fullscreen change, and device-pixel-ratio change. A red `⚠ TEST PATTERN IS NOT BEING DISPLAYED 1:1` banner appears whenever it isn't
- **MT Vertical Content Mapping: 2:1** (MT only) - a separate, static fact about MT's content-to-physical-LED relationship, kept deliberately distinct from the browser mapping check above so the two can never be confused for one another
- **Fit to Output** - only shown when the test pattern is larger than the display - deliberately scales the canvas to fill the display's width (aspect ratio preserved) instead of overflowing/clipping. Clearly labeled "Scaled to fit output" (a calm, blue message) rather than the red not-1:1 warning, since this is a choice you made, not an accidental mismatch. "Show Native Resolution (1:1)" switches back

When you open the Moving Test Pattern, this app also tries to automatically open and fullscreen it on a second monitor if one's connected, using your browser's Window Management API (Chrome/Edge, over HTTPS or `localhost`, with permission) - the main app window stays where it is. With more than one secondary display connected, it asks which one to use and remembers the choice; "Change output display" (next to the button) forgets that choice so it asks again next time. If your browser doesn't support this (Firefox/Safari, an insecure context, or permission not granted), it falls back to opening a plain window - move and fullscreen that one manually - with a one-time note explaining why. If the automatic fullscreen attempt is blocked by the browser (common for a window opened programmatically, without your own direct click inside it), a large **ENTER FULLSCREEN** button appears in that window - click it to finish the job yourself.

## Rentman Integration

Optional. Lets Stock Calculations pull live on-hand stock counts, date-range availability and broken/under-repair quantities from [Rentman](https://www.rentman.io/). This app is a static site with no backend, and the Rentman API token must never end up in anything shipped to the browser - so this works through a small separate Cloudflare Worker that holds the token server-side and proxies read-only requests to Rentman.

Setup (one-time):

1. Deploy the Worker - see [`rentman-proxy/README.md`](./rentman-proxy/README.md) for the full steps (`wrangler secret put`, `wrangler deploy`).
2. Copy [`.env.example`](./.env.example) to `.env` for local dev, and/or add a `RENTMAN_PROXY_URL` repository **variable** (not secret) under `Settings -> Secrets and variables -> Actions -> Variables`) so the GitHub Pages build picks it up via [`deploy-pages.yml`](./.github/workflows/deploy-pages.yml).
3. Rebuild/redeploy. The Rentman controls at the top of **Stock Calculations** will show as configured; click **Get Current Stock from Rentman**, **Check Stock Availability by Date Range** or **Check Broken / Repair Equipment** - no mapping step needed, since this catalog's codes already match Rentman's own equipment codes.

Left unset, that block just shows "Not configured" and Stock Calculations keeps using its built-in numbers - nothing else changes.

Availability arithmetic, in one line:

```
Available Stock = Rentman Stock - Other Projects (in the date range) - Broken / Repair
```

checked against this project's own **Total Required** (required panels plus the box-rounded spare). Broken/repair is deliberately *not* date-ranged - equipment in the workshop is off the shelf today, whenever the job is.
