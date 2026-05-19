import { FN001001_DEP_microsoft_sp_core_library } from './rules/FN001001_DEP_microsoft_sp_core_library';
import { FN001002_DEP_microsoft_sp_lodash_subset } from './rules/FN001002_DEP_microsoft_sp_lodash_subset';
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from './rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core';
import { FN001004_DEP_microsoft_sp_webpart_base } from './rules/FN001004_DEP_microsoft_sp_webpart_base';
import { FN001011_DEP_microsoft_sp_dialog } from './rules/FN001011_DEP_microsoft_sp_dialog';
import { FN001012_DEP_microsoft_sp_application_base } from './rules/FN001012_DEP_microsoft_sp_application_base';
import { FN001013_DEP_microsoft_decorators } from './rules/FN001013_DEP_microsoft_decorators';
import { FN001014_DEP_microsoft_sp_listview_extensibility } from './rules/FN001014_DEP_microsoft_sp_listview_extensibility';
import { FN001021_DEP_microsoft_sp_property_pane } from './rules/FN001021_DEP_microsoft_sp_property_pane';
import { FN001023_DEP_microsoft_sp_component_base } from './rules/FN001023_DEP_microsoft_sp_component_base';
import { FN001024_DEP_microsoft_sp_diagnostics } from './rules/FN001024_DEP_microsoft_sp_diagnostics';
import { FN001025_DEP_microsoft_sp_dynamic_data } from './rules/FN001025_DEP_microsoft_sp_dynamic_data';
import { FN001026_DEP_microsoft_sp_extension_base } from './rules/FN001026_DEP_microsoft_sp_extension_base';
import { FN001027_DEP_microsoft_sp_http } from './rules/FN001027_DEP_microsoft_sp_http';
import { FN001028_DEP_microsoft_sp_list_subscription } from './rules/FN001028_DEP_microsoft_sp_list_subscription';
import { FN001029_DEP_microsoft_sp_loader } from './rules/FN001029_DEP_microsoft_sp_loader';
import { FN001030_DEP_microsoft_sp_module_interfaces } from './rules/FN001030_DEP_microsoft_sp_module_interfaces';
import { FN001031_DEP_microsoft_sp_odata_types } from './rules/FN001031_DEP_microsoft_sp_odata_types';
import { FN001032_DEP_microsoft_sp_page_context } from './rules/FN001032_DEP_microsoft_sp_page_context';
import { FN001034_DEP_microsoft_sp_adaptive_card_extension_base } from './rules/FN001034_DEP_microsoft_sp_adaptive_card_extension_base';
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from './rules/FN002002_DEVDEP_microsoft_sp_module_interfaces';
import { FN002021_DEVDEP_rushstack_eslint_config } from './rules/FN002021_DEVDEP_rushstack_eslint_config';
import { FN002022_DEVDEP_microsoft_eslint_plugin_spfx } from './rules/FN002022_DEVDEP_microsoft_eslint_plugin_spfx';
import { FN002023_DEVDEP_microsoft_eslint_config_spfx } from './rules/FN002023_DEVDEP_microsoft_eslint_config_spfx';
import { FN002024_DEVDEP_eslint } from './rules/FN002024_DEVDEP_eslint';
import { FN002025_DEVDEP_eslint_plugin_react_hooks } from './rules/FN002025_DEVDEP_eslint_plugin_react_hooks';
import { FN002030_DEVDEP_microsoft_spfx_web_build_rig } from './rules/FN002030_DEVDEP_microsoft_spfx_web_build_rig';
import { FN002031_DEVDEP_rushstack_heft } from './rules/FN002031_DEVDEP_rushstack_heft';
import { FN002032_DEVDEP_typescript_eslint_parser } from './rules/FN002032_DEVDEP_typescript_eslint_parser';
import { FN002034_DEVDEP_microsoft_spfx_heft_plugins } from './rules/FN002034_DEVDEP_microsoft_spfx_heft_plugins';
import { FN002035_DEVDEP_types_heft_jest } from './rules/FN002035_DEVDEP_types_heft_jest';
import { FN002036_DEVDEP_types_jest } from './rules/FN002036_DEVDEP_types_jest';
import { FN010001_YORC_version } from './rules/FN010001_YORC_version';
import { FN015008_FILE_eslintrc_js } from './rules/FN015008_FILE_eslintrc_js';
import { FN015016_FILE_eslint_config_js } from './rules/FN015016_FILE_eslint_config_js';
import { FN022001_SCSS_remove_fabric_react } from './rules/FN022001_SCSS_remove_fabric_react';
import { FN022002_SCSS_add_fabric_react } from './rules/FN022002_SCSS_add_fabric_react';
import { FN027001_OVERRIDES_rushstack_heft } from './rules/FN027001_OVERRIDES_rushstack_heft';

module.exports = [
  new FN001001_DEP_microsoft_sp_core_library('1.23.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.23.0'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.23.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.23.0'),
  new FN001011_DEP_microsoft_sp_dialog('1.23.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.23.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.23.0'),
  new FN001021_DEP_microsoft_sp_property_pane('1.23.0'),
  new FN001023_DEP_microsoft_sp_component_base('1.23.0'),
  new FN001024_DEP_microsoft_sp_diagnostics('1.23.0'),
  new FN001025_DEP_microsoft_sp_dynamic_data('1.23.0'),
  new FN001026_DEP_microsoft_sp_extension_base('1.23.0'),
  new FN001027_DEP_microsoft_sp_http('1.23.0'),
  new FN001028_DEP_microsoft_sp_list_subscription('1.23.0'),
  new FN001029_DEP_microsoft_sp_loader('1.23.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.23.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.23.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.23.0'),
  new FN001013_DEP_microsoft_decorators('1.23.0'),
  new FN001034_DEP_microsoft_sp_adaptive_card_extension_base('1.23.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.23.0'),
  new FN002022_DEVDEP_microsoft_eslint_plugin_spfx('1.23.0'),
  new FN002023_DEVDEP_microsoft_eslint_config_spfx('1.23.0'),
  new FN002030_DEVDEP_microsoft_spfx_web_build_rig('1.23.0'),
  new FN002034_DEVDEP_microsoft_spfx_heft_plugins('1.23.0'),
  new FN010001_YORC_version('1.23.0'),
  new FN002031_DEVDEP_rushstack_heft('1.2.17'),
  new FN027001_OVERRIDES_rushstack_heft('1.2.17'),
  new FN002025_DEVDEP_eslint_plugin_react_hooks('5.2.0'),
  new FN002024_DEVDEP_eslint('9.37.0'),
  new FN015016_FILE_eslint_config_js(true, `const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    }
  }
];`),
  new FN015008_FILE_eslintrc_js(false),
  new FN002021_DEVDEP_rushstack_eslint_config('4.5.2', false),
  new FN002032_DEVDEP_typescript_eslint_parser('8.46.2', false),
  new FN002035_DEVDEP_types_heft_jest('1.0.2', false),
  new FN002036_DEVDEP_types_jest('30.0.0'),
  new FN022001_SCSS_remove_fabric_react('~@fluentui/react/dist/sass/References.scss'),
  new FN022002_SCSS_add_fabric_react('pkg:@fluentui/react/dist/sass/References.scss')
];