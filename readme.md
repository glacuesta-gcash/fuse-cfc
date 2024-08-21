# Cascading Forecasts

This is the **Cascading Forecasts (CFC)** system. It provides an automated solution for the efficient management and use of base models, for concurrent, overlapping forecasts. 

While originally developed with financial forecasting in mind (in relation to annual budgeting exercises), this is a flexible and scalable system that can be used for any modeling application. It should in particular be helpful in cases where a single base model is used for multiple products or lines of business. By automating and aiding in the duplication and setup of the base model, as well as the gathering of outputs into a single summary, this simplifies forecasting by only needing the user to set up the base model once, and then just focus on the assumptions afterwards.

## Pain Points

- Changing the calculation models required the user to reapply the changes to all calculation tabs, a very laborious and error-prone proposition
- The whole thing is prone to typo errors when mapping and dragging cell references from assumption tabs to calc tabs, and calc tabs to summary tabs
- Very easy to forget the overall story when trying to summarize the assumptions and forecast, because overrides and inputs are all over the whole file
- Running scenarios / alternatives required plenty of manual copy-pasting, reapplication or changing of values scattered all over the spreadsheet

## Objectives

- Forecasting done procedurally:
    - Parametric modeling of input assumptions
    - Replicable results given the same generation steps and templates
- Single place (tab templates) for formulas to be defined and set up
- Discourage/prevent overriding of arbitrary ranges inside the projected periods
- Allow easy setup and comparison of multiple scenarios, for high/mid/low building, sensitivity studies, etc.

## Sheet Tabs

There should be a `_steps` tab.

There should also be a `_summary` tab.

Tabs preceded with `_` (underscore) indicate input tabs that can be used as templates to be duplicated.

Tabs preceded with `-` (hyphen) indicate tabs generated by CFC (dynamic tabs) and will be cleaned up on execution.

All other tabs will be ignored.

## Inside a tab

#### Period column
The first row is reserved for special designations and labels. It may have 1 cell labeled `p` (the "period column") which shall be the one duplicated in order to fill out the forecast periods. This column can contain formulas.

<details>
<summary>Example</summary>

#### Tab `_purchases`
        A           B           C
    1               p
    2   base        1000
    3   takeup-rate 3%
    4   takeup      =B2*B3

</details>

#### Other columns
There may also be other columns with other labels in the first row, such as "p0", "consumer-low", or any other arbitrary text. These can be useful for setting up initialization columns, to the left of the period column, which are set up with static values, while the period columns follow a different formula. This can also be used in assumption tabs to have parallel sets of variables for different use cases, target markets, etc.

<details>
<summary>Example</summary>

#### Tab `_assumptions`

        A           B           C
    1               consumer    merchant
    2   takeup-rate 5%          3.2%
    3   churn-rate  14%         8%

Later on, `map` can be used to set variables in forecast tabs:

`map assumptions/takeup-rate:consumer purchases-consumer/takeup-rate`
`map assumptions/takeup-rate:merchant purchases-merchant/takeup-rate`
</details>

#### "A" Column - Reserved for variable names
The first column (A) should only contain variable names. A variable name can end with `:x` to indicate a variable that spans multiple rows, where `x` indicates the total number of rows that variable covers.

<details>
<summary>Example</summary>

#### Tab `_members`

        A           B           C           D           E
    1                           p0          p
    2   members     Members     1000        =C2+100
    3   mem_fee     Fee/member              100
    4   fees        Total Fees              =D2*D3

#### Tab `_assumptions`

        A           B           C           D           E
    1               scenario1   scenario2
    2   start_mems  1000        1500

> In the above example, `members:p0` can be initialized to other values via a `map assumptions start_mems:scenario1 members members:p0` command, and each period afterwards (`p1` onwards) will increment this by 100, given the formula in D2.

> The `mem_fee` variable, meanwhile, can be trended over time via `trend members mem_fee p1 p12 100 200 linear` to increase the membership fee each month.
</details>

## Step Reference

### Setup Commands

#### Set
`set [setting] [value]`

Set a configuration setting to a given value. The following are the configurable
 settings:

- `periods` - The number of periods (e.g. months) to forecast. Defaults to 12.
- `summary-periods` - The number of periods inside 1 summary column. Defaults to 12 (1 year). The summary will include as many summary periods as can fit inside the total forecast. Using "1" will not have grouped columns.
- `summary-start` - The starting period (e.g. `p1`, `p6`) for the summary. For instance, starting at `p6` with 12 summary periods will summarize `p6` to `p17` (12 periods) as the first summary column. Use this if, for instance, you are starting the forecast from September but want to summarize full years starting the following January.

For example, if the forecast starts July of Year 0, and you want quarterly summaries, for 2 years starting January of Year 1:

    set     periods             30
    set     summary-periods     3
    set     summary-start       6

#### Build
`build [source tab]`

Build out an assumption tab. A new tab will be spawned, to keep the original pristine, and the period column will be duplicated to extend its forecast to the appropriate number of periods.

#### Spawn
`spawn [source tab] [new tab]`
`spawn [source tab] [new tab],[new tab]...`

Copies the tab `[source tab]` to `[new tab]`. The new tab will be given the prefix `-` (hyphen) to mark it as a generated tab, and will be deleted on next execution.

More than one new tab can be specified, by delimiting multiple names with a comma. For example:

    spawn   members   mem-corp,mem-pers,mem-free

`build` will also be performed on the spawned tab(s).

---

### Value initialization and setting commands

<details>
<summary>Note: Multi-row variables</summary>
&nbsp;

> If a variable in the tab follows the pattern `var_name|n` where *n* is a number, this indicates that the variable is to be treated as a stack of *n* values. Thus, the `map` command will operate on all *n* rows of that variable.

        A           B           C           D           E
    1                           p0          p
    2   members     Members     1000        =C2+100
    3   age_grps|4  18-30       40%         =D$2*$C3
    4               31-40       30%         =D$2*$C4
    5               41-50       20%         =D$2*$C5
    6               51+         10%         =D$2*$C6
    7               Txns/mo
    8   age_txns|4  18-30                   4.2
    9               31-40                   3.7
    10              41-50                   3.1
    11              51+                     2.8
    12              Total txns              =SUMPRODUCT(D3:D6,D8:D11)

>In the example above, referencing `age_grps` will apply to rows 3-6, and `age_txns` will apply to rows 8-11.

>This is limited to the `map` command (for now).

</details>

#### Bump
`bump [tab]/[var] [start period] [val]`

Sets the value of the variable to val, starting at `[start period]` all the way to the end of the forecast.

> Note: Bump cannot be performed on a multi-row variable

#### Trend
`trend [tab]/[var] [start period] [end period] [start val] [end val] *[method]`

Apply a trend (e.g. growth over time) onto a variable in a given tab. The variable will be blended from `[start val]` in `[start period]` to `[end val]` in `[end period]`, depending on the `[method]` (defaults to linear if omitted):
- `linear` - Straight linear growth
- `expo` - Use a fixed periodic growth rate

> Note: Trend cannot be performed on a multi-row variable

#### Map
`map [source tab] [source var] [target tab] [target var]`

Maps one variable from one tab (the source) to another variable in another tab (the target). The mapping is done via cell reference, to make it easier for others to trace the logic in the worksheet.

`map [source tab] [source var]:[source col] [target tab] [target var]:[target col]`

Similar to the above pattern, although here, source and target columns are specified. Given this, only 1 value will be mapped, instead of the entire time horizon.

---

### Cleanup and Summary Commands

#### Summarize

`summarize [var] [method]`

Include a variable in the summary:
* In the Summary tab, the row with the variable specified (in col A) will be duplicated as needed, in order to populate each row with the mapped values from all dynamic tabs that also contain the variable. (For instance, a row with `gross-rev` will be used as the template for capturing the `gross-rev` variable in all applicable dynamic tabs.)
* If a row is found with the variable specified, suffixed with a `!` (exclamation point) will be populated with the sum of that variable, for each period. (For instance, a row with `gross-rev!` in col A.) Let's call this the vertical summary. It will always be a sum.
* Periodic column groups (horizontal summaries) will also be added (e.g. semestral or annual summaries), using `method` (`sum`, `average`, or `last`) to calculate the appropriate time period summary value

#### Group

(to do)

`group [label] [tab],[tab],[tab]...`

Combine tabs into a group when summarizing. Tabs should be comma-separated. Example:

    group gloan gloan-c-low,gloan-c-med,gloan-c-high

This will cause the three tabs to be listed and subtotaled under the label `gloan` in the Summary tab, where applicable.

#### Scenario

(to do)

`scenario [scenario]`

Causes all the following steps (until another `scenario` step, or the end of all steps) to be considered to be only under this specific scenarios.

If using scenarios, one Summary tab will be generated for each scenario, labeled `-summary-[scenario]`.

### To-do

- Copying/mapping from preexisting period cells (i.e. monthly input assumptions)? Is this needed though, given you can use vanilla worksheet functions to pre-map from p col.
- ✔ Order of tabs (i.e. in spawn) should be guaranteed despite parallel/batch duplication
- ✔ Grouping and subtotaling in summary
- ✔ Friendly titles for tabs
- ✔ Group periods, including last, sum, average
- ✔ Multirow vars
- ✔ :col specification in commands (map)
- ✔ Asynchronous threads for combining read and write ops
- ✔ Batch update for spawn
- ✔ Group raw period columns in summary

### Other notes

- Where applicable, threading and batched updates are used to optimize Google API calls, to significantly bring down execution time.

---
2024 G Lacuesta