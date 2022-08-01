import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import ProductModelDialog from './components/PageModelDialog';
import { setup as pnpSetup } from "@pnp/common";
import { Dialog } from '@microsoft/sp-dialog';



/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface BBS_CPA_IpageCommandSetCommandSetProperties {
  // This is an example; replace with your own properties
}

const LOG_SOURCE: string = 'pageCommandSetCommandSet';

export default class BBS_CPA_pageCommandSetCommandSet extends BaseListViewCommandSet<BBS_CPA_IpageCommandSetCommandSetProperties> {

  @override
  public onInit(): Promise<void> {

    Log.info(LOG_SOURCE, 'Initialized pageCommandSetCommandSet');

    return super.onInit().then(_ => {

      // other init code may be present

      pnpSetup({
        spfxContext: this.context
      });
    });
  }



  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {

  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {

    switch (event.itemId) {
      case 'COMMAND_1':
        const dialog: ProductModelDialog = new ProductModelDialog(this.context.pageContext.web.absoluteUrl);
        dialog.show();

        break;
        
      default:
        throw new Error('Unknown command');

    }
  }

}
