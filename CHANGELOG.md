# Change Log

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

