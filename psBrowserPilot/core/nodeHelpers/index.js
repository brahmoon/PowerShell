import * as NodeUI from './NodeUI.js';
import { createNodeBridge } from './NodeBridge.js';
import * as ExcelHelper from './ExcelHelper.js';
import * as ColorHelper from './ColorHelper.js';
import { createPaletteController } from './PaletteUI.js';
import { createStorageHelper } from './StorageHelper.js';
import { createNodeLogger } from './NodeLogger.js';
import * as UIHelper from './UIHelper.js';

export const createHelperContext = (baseContext = {}) => {
  const logger = createNodeLogger(baseContext);
  const storage = createStorageHelper(baseContext);
  const bridge = createNodeBridge(baseContext, { logger });

  return {
    NodeUI,
    NodeBridge: bridge,
    ExcelHelper,
    ColorHelper,
    PaletteUI: {
      create: (options = {}) => createPaletteController({ ...options, storage }),
    },
    StorageHelper: storage,
    NodeLogger: logger,
    UIHelper,
  };
};

export { NodeUI };
export { createNodeBridge, DEFAULT_SERVER_URL, RUN_SCRIPT_PATH } from './NodeBridge.js';
export { ExcelHelper };
export { ColorHelper };
export { createPaletteController } from './PaletteUI.js';
export { createStorageHelper } from './StorageHelper.js';
export { createNodeLogger } from './NodeLogger.js';
export { UIHelper };
