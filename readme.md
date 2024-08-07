# Cascading Forecasts

### Objectives

- Forecasting done procedurally:
    - Parametric modeling of input assumptions
    - Replicable results given the same generation steps and templates
- Single place (tab templates) for formulas to be defined and set up
- Discourage/prevent overriding of arbitrary ranges inside the projected periods
- Allow easy setup and comparison of multiple scenarios, for high/mid/low building, sensitivity studies, etc.

### Sheet Tabs

Tabs preceded with `_` (underscore) indicate calculation tabs that can be used as templates to be duplicated.

There should be a `_steps` tab.

Tabs preceded with `-` (hyphen) indicate tabs generated by CFC and will be cleaned up on execution.

All other tabs will be ignored.

### Inside a tab

The first row (1) should be blank, except for 1 cell labeled "p" which shall be the one duplicated in order to fill out the forecast periods. This column can contain formulas.

The first column (A) should only contain variable names. A variable name can end with `:x` to indicate a variable that spans multiple rows, where `x` indicates the total number of rows that variable covers.

### Commands

#### Set
`set [setting] [value]`
Set a configuration setting to a given value. The following are the configurable
 settings:

- `periods` - The number of periods (e.g. months) to forecast. Defaults to 12.

#### Build
`build [source tab]`

Build out an assumption tab. A new tab will be spawned, to keep the original pristine, and the period column will be duplicated to extend its forecast to the appropriate number of periods.

#### Spawn
`spawn [source tab] [new tab]`

Copies the tab `[source tab]` to `[new tab]`. The new tab will be given the prefix `-` (hyphen) to mark it as a generated tab, and will be deleted on next execution.

`build` will also be performed on the spawned tab.

#### Bump
`bump [tab] [var] [start period] [val]`

Sets the value of the variable to val, starting at `[start period]` all the way to the end of the forecast.

#### Trend
`trend [tab] [var] [start period] [end period] [start val] [end val] *[method]`

Apply a trend (e.g. growth over time) onto a variable in a given tab. The variable will be blended from `[start val]` in `[start period]` to `[end val]` in `[end period]`, depending on the `[method]` (defaults to linear if omitted):
- `linear` - Straight linear growth
- `expo` - Use a fixed periodic growth rate

#### Map
`map [source tab] [source var] [target tab] [target var]`

Maps one variable from one tab (the source) to another variable in another tab (the target). The mapping is done via cell reference, to make it easier for others to trace the logic in the worksheet.