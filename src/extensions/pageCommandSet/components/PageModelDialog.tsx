import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';

import ProductModelDialogContent from './PageModelDialogContent';
/**
 * class permet de définir les évènements de la Boite de dialogue lancé par la commande
 */
export default class BBS_CPA_ProductModelDialog extends BaseDialog {
 
  //Rendre le Contenu de la boite de dialogue avec le component react 
  public render(): void {
    
    ReactDOM.render(<ProductModelDialogContent 
      close={ this.close.bind(this) } 

    />, this.domElement);

  }
 
  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }
  //fermer la boite de dialogue
  public close(): Promise<void> {

    return super.close();

  }
  //unmount react component 
  protected onAfterClose(): void {
    super.onAfterClose();
    super.close();
    // Clean up the element for the next dialog
    ReactDOM.unmountComponentAtNode(this.domElement);
    
  }
 


}

