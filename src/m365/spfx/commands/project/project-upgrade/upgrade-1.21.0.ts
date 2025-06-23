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
import { FN002001_DEVDEP_microsoft_sp_build_web } from './rules/FN002001_DEVDEP_microsoft_sp_build_web';
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from './rules/FN002002_DEVDEP_microsoft_sp_module_interfaces';
import { FN002022_DEVDEP_microsoft_eslint_plugin_spfx } from './rules/FN002022_DEVDEP_microsoft_eslint_plugin_spfx';
import { FN002023_DEVDEP_microsoft_eslint_config_spfx } from './rules/FN002023_DEVDEP_microsoft_eslint_config_spfx';
import { FN002024_DEVDEP_eslint } from './rules/FN002024_DEVDEP_eslint';
import { FN002026_DEVDEP_typescript } from './rules/FN002026_DEVDEP_typescript';
import { FN002029_DEVDEP_microsoft_rush_stack_compiler_5_3 } from './rules/FN002029_DEVDEP_microsoft_rush_stack_compiler_5_3';
import { FN010001_YORC_version } from './rules/FN010001_YORC_version';
import { FN012017_TSC_extends } from './rules/FN012017_TSC_extends';
import { FN021003_PKG_engines_node } from './rules/FN021003_PKG_engines_node';

module.exports = [
  new FN001001_DEP_microsoft_sp_core_library('1.21.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.21.0'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.21.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.21.0'),
  new FN001011_DEP_microsoft_sp_dialog('1.21.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.21.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.21.0'),
  new FN001021_DEP_microsoft_sp_property_pane('1.21.0'),
  new FN001023_DEP_microsoft_sp_component_base('1.21.0'),
  new FN001024_DEP_microsoft_sp_diagnostics('1.21.0'),
  new FN001025_DEP_microsoft_sp_dynamic_data('1.21.0'),
  new FN001026_DEP_microsoft_sp_extension_base('1.21.0'),
  new FN001027_DEP_microsoft_sp_http('1.21.0'),
  new FN001028_DEP_microsoft_sp_list_subscription('1.21.0'),
  new FN001029_DEP_microsoft_sp_loader('1.21.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.21.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.21.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.21.0'),
  new FN001013_DEP_microsoft_decorators('1.21.0'),
  new FN001034_DEP_microsoft_sp_adaptive_card_extension_base('1.21.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.21.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.21.0'),
  new FN002024_DEVDEP_eslint('8.57.1'),
  new FN002022_DEVDEP_microsoft_eslint_plugin_spfx('1.21.0'),
  new FN002023_DEVDEP_microsoft_eslint_config_spfx('1.21.0'),
  new FN002026_DEVDEP_typescript('5.3.3'),
  new FN002029_DEVDEP_microsoft_rush_stack_compiler_5_3('0.1.0'),
  new FN010001_YORC_version('1.21.0'),
  new FN012017_TSC_extends('./node_modules/@microsoft/rush-stack-compiler-5.3/includes/tsconfig-web.json'),
  new FN021003_PKG_engines_node('>=22.14.0 < 23.0.0')
];
