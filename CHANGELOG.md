# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.71] - 2019-12-24

### Fixed
- bug when air source changed to `inside` after using `send()` function of `Breezer` in manual mode
### Added
- `Breezer` parameter `gate` (air source) now can be controled in `manual` zone mode

## [0.6] - 2019-12-02

### Changed
- `Breezer` parameter `is_on` now can't be set mannualy and calculated automatically depending on the speed
- all parameters of `Zone`, `MagicAir` and `Breezer`, that can't be changed, became `@property`

## [0.5] - 2019-12-01

### Added
- `min_update_interval` option for TionApi class
- `force` parameter for load() functions, which allows to get new data immediately regardless `min_update_interval` option
- verification for `zone_data` and `device_data` in `load()` functions to avoid exceptions if Tion server goes offline
- Breezer parameters:
  - heater_installed
  - t_min
  - t_max
  - speed_limit
- usage example as main() function

### Changed
- print replaced by `_LOGGER`
- `zone` objects in `Breezer` and `MagicAir` classes now `Zone` objects, not just raw data
- headers in `TionApi` now property, to update authorization after init
- new tests made from scratch
- canceled checking `__repr__()` and other hardly reachable places by coverage in tests
