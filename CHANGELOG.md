# Change Log

## [2.2.0] - 2023-02-12
### Added
- An `excel` macro to capture output from `print-table`. See README.

## [2.1.0] - 2022-02-21
### Changed
- Update dependencies and resolve JODConverter breaking changes, including 
  changes which mitigate some vulnerabilities in commons-compress and 
  commons-io. See [#13](https://github.com/matthewdowney/excel-clj/pull/13).


## [2.0.1] - 2021-02-22

### Changed
 - Updated dependency versions for taoensso/encore and taoensso/tufte
### Added
- Support for `LocalDate` and `LocalDateTime` (see 
  [#9](https://github.com/matthewdowney/excel-clj/pull/9)).

## [2.0.0] - 2020-10-04
### Changed
- Now uses the POI streaming writer by default (~10x performance gain on 
  sheets > 100k rows)
- Separated out writer abstractions in [poi.clj](src/excel_clj/poi.clj) to 
  allow using a lower-level POI interface
- Simplified & rewrote [tree.clj](src/excel_clj/tree.clj)
- Better wrapping for styling and dimension data in 
  [cell.clj](src/excel_clj/cell.clj)

### Added 
- Support for merging workbooks, so you can have a template which uses formulas
  which act on data from some named sheet, and then fill in that named sheet.
- New top-level helpers for working with grid (`[[cell]]`) data structures
- Vertical as well as horizontal merged cells
- New constructors to build grids from tables and trees (`table-grid` and 
  `tree-grid`), which supplant the deprecated constructors from v1.x (`tree` 
  and `table`)
  
## [1.3.3] - 2020-07-11
### Fixed
- Bug where columns would only auto resize up until 'J'
- Unnecessary Rhizome dependency causing headaches in headless environments

## [1.3.2] - 2020-04-15
### Fixed
- Bug introduced in v1.3.1 where adjacent cells with width > 1 cause an 
  exception.

## [1.3.1] - 2020-04-05
### Added
- A lower-level, writer style interface for Apache POI.
- [Prototype/brainstorm](src/excel_clj/prototype.clj) of less complicated, 
  pure-data replacement for high-level API in upcoming v2 release.
### Fixed
- Bug (#3) with the way cells were being written via POI that would write cells
  out of order or mix up the style data between cells.

## [1.2.1] - 2020-04-01
### Added
- Can bind a dynamic `*n-threads*` var to set the number of threads used during 
  writing.

## [1.2.0] - 2020-08-13
### Added
- Performance improvements for large worksheets.

## [1.1.2] - 2019-06-04
### Fixed
- If the first level of the tree is a leaf, `accounting-table` doesn't walk it 
  correctly.
### Added
- Can pass through a custom `:min-leaf-depth` key to `tree` (replaces binding a 
dynamic var in earlier versions).

## [1.1.1] - 2019-06-01
### Fixed
- Total rows were not always being displayed correctly for trees

## [1.1.0] - 2019-05-28
### Added
- More flexible tree rendering/aggregation

### Changed
- Replaced lots of redundant tree code with a `walk` function

## [1.0.0] - 2019-04-17
### Added
- PDF generation
- Nicer readme, roadmap & tests

## [0.1.0] - 2019-01-15
- Pulled this code out of an accounting project I was working on as its own library.
- Already had
    - Clojure data wrapper over Apache POI.
    - Tree/table/cell specifications.
    - Excel sheet writing.

