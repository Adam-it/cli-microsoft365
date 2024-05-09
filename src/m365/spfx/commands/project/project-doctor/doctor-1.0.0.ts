import { FN002004_DEVDEP_gulp } from './rules/FN002004_DEVDEP_gulp';
import { FN002013_DEVDEP_types_webpack_env } from './rules/FN002013_DEVDEP_types_webpack_env';
import { FN021001_PKG_spfx_deps_versions_match_project_version } from './rules/FN021001_PKG_spfx_deps_versions_match_project_version';

module.exports = [
  new FN002004_DEVDEP_gulp('~3.9.1'),
  new FN002013_DEVDEP_types_webpack_env('>=1.12.1 <1.14.0'),
  new FN021001_PKG_spfx_deps_versions_match_project_version()
];