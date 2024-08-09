# Cascading Forecasts

This is the Cascading Forecasts (CFC) system.

### Pain Points

- Changing the calculation models required the user to reapply the changes to all calculation tabs, a very laborious and error-prone proposition
- The whole thing is prone to typo errors when mapping and dragging cell references from assumption tabs to calc tabs, and calc tabs to summary tabs
- Very easy to forget the overall story when trying to summarize the assumptions and forecast, because overrides and inputs are all over the whole file
- Running scenarios / alternatives required plenty of manual copy-pasting, reapplication or changing of values scattered all over the spreadsheet

### Objectives

- Forecasting done procedurally:
    - Parametric modeling of input assumptions
    - Replicable results given the same generation steps and templates
- Single place (tab templates) for formulas to be defined and set up
- Discourage/prevent overriding of arbitrary ranges inside the projected periods
- Allow easy setup and comparison of multiple scenarios, for high/mid/low building, sensitivity studies, etc.

### Sheet Tabs

There should be a `_steps` tab.

There should also be a `_summary` tab.

Tabs preceded with `_` (underscore) indicate input tabs that can be used as templates to be duplicated.

Tabs preceded with `-` (hyphen) indicate tabs generated by CFC (dynamic tabs) and will be cleaned up on execution.

All other tabs will be ignored.

### Inside a tab

The first row (1) should be blank, except for 1 cell labeled "p" which shall be the one duplicated in order to fill out the forecast periods. This column can contain formulas.

The first column (A) should only contain variable names. A variable name can end with `:x` to indicate a variable that spans multiple rows, where `x` indicates the total number of rows that variable covers.

### Step Reference

#### Set
`set [setting] [value]`

Set a configuration setting to a given value. The following are the configurable
 settings:

- `periods` - The number of periods (e.g. months) to forecast. Defaults to 12.
- `summary-periods` - The number of periods inside 1 summary column. Defaults to 12 (1 year). The summary will include as many summary periods as can fit inside the total forecast. Using "1" will not have grouped columns.
- `summary-start` - The starting period (e.g. `p1`, `p6`) for the summary. For instance, starting at `p6` with 12 summary periods will summarize `p6` to `p17` (12 periods) as the first summary column. Use this if, for instance, you are starting the forecast from September but want to summarize full years starting the following January.

#### Build
`build [source tab]`

Build out an assumption tab. A new tab will be spawned, to keep the original pristine, and the period column will be duplicated to extend its forecast to the appropriate number of periods.

Cached update.

#### Spawn
`spawn [source tab] [new tab]`

Copies the tab `[source tab]` to `[new tab]`. The new tab will be given the prefix `-` (hyphen) to mark it as a generated tab, and will be deleted on next execution.

`build` will also be performed on the spawned tab.

Cached update.

#### Bump
`bump [tab] [var] [start period] [val]`

Sets the value of the variable to val, starting at `[start period]` all the way to the end of the forecast.

Cached update.

#### Trend
`trend [tab] [var] [start period] [end period] [start val] [end val] *[method]`

Apply a trend (e.g. growth over time) onto a variable in a given tab. The variable will be blended from `[start val]` in `[start period]` to `[end val]` in `[end period]`, depending on the `[method]` (defaults to linear if omitted):
- `linear` - Straight linear growth
- `expo` - Use a fixed periodic growth rate

Cached update.

#### Map
`map [source tab] [source var] [target tab] [target var]`

Maps one variable from one tab (the source) to another variable in another tab (the target). The mapping is done via cell reference, to make it easier for others to trace the logic in the worksheet.

Cached update.

#### Summarize

(to do)

`summarize [var] [method]`

Include a variable in the summary:
* In the Summary tab, the row with the variable specified (in col A) will be duplicated as needed, in order to populate each row with the mapped values from all dynamic tabs that also contain the variable. (For instance, a row with `gross-rev` will be used as the template for capturing the `gross-rev` variable in all applicable dynamic tabs.)
* If a row is found with the variable specified, suffixed with a `!` (exclamation point) will be populated with the sum of that variable, for each period. (For instance, a row with `gross-rev!` in col A.) Let's call this the vertical summary. It will always be a sum.
* Periodic column groups (horizontal summaries) will also be added (e.g. semestral or annual summaries), using `method` (`sum`, `average`, or `last`) to calculate the appropriate time period summary value

#### Group

(to do)

`group [label] [tabs]`

Combine tabs into a group when summarizing. Tabs should be comma-separated. Example:

    group gloan gloan-c-low,gloan-c-med,gloan-c-high

This will cause the three tabs to be listed and subtotaled under the label `gloan` in the Summary tab, where applicable.

#### Scenario

(to do)

`scenario [scenario]`

Causes all the following steps (until another `scenario` step, or the end of all steps) to be considered to be only under this specific scenarios.

If using scenarios, one Summary tab will be generated for each scenario, labeled `-summary-[scenario]`.

---
2024 G Lacuesta