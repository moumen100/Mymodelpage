import * as React from 'react';
import {
  Dialog,
  DialogFooter,
  DefaultButton,
  PrimaryButton,
  ChoiceGroup,
  DialogType,
  TextField,
  Dropdown,
  Spinner,
  SpinnerSize,
  Label
}
  from 'office-ui-fabric-react';

import { ProductModelHelper } from './PageModelHelper';
import { sp as sp } from "@pnp/sp";
import "./PageModelDialog.module.scss";

export interface BBS_CPA_IProductModelDialogContentState {
  hideDialog: boolean;
  optionSelected: string;
  newPageName: string;
  selectedItem?: { key: string | number | undefined };
  selectedItems: string[];
  allItems: any[];
  isLoading: boolean;
  creationDone: boolean;
}

export interface BBS_CPA_IProductModelDialogContentProps {
  close: () => void;
}


export default class BBS_CPA_ProductModelDialogContent extends React.Component<BBS_CPA_IProductModelDialogContentProps, BBS_CPA_IProductModelDialogContentState> {

  constructor(props: BBS_CPA_IProductModelDialogContentProps) {
    super(props);
    this.state = {
      hideDialog: false,
      optionSelected: 'A',
      newPageName: "",
      selectedItem: null,
      selectedItems: null,
      allItems: null,
      isLoading: false,
      creationDone: false
    };

    sp.web.lists.getByTitle("Site Pages").items.filter("Model eq 'Product'").select("Title,FileRef").getAll().then((items: any[]) => {
      var tmpItems: any[] = new Array();
      //DropDown initialization
      items.forEach(element => {
        var item = { key: element["FileRef"], text: element["Title"] };
        tmpItems.push(item);
      });
      console.log("Checking values ----");
        console.log(tmpItems);
      this.setState({ allItems: tmpItems });
    });
   
  }


  public render() {

    return (
      <div>
        <Dialog
          hidden={this.state.hideDialog}

          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: "Créer un nouveau produit à partir d'une template"
          }}

          modalProps={{
            isBlocking: false,
            containerClassName: 'ms-dialogMainOverride'
          }}
        >
          {!this.state.creationDone && (
            <div>
              <ChoiceGroup
                label="Choisir un type"
                options={[
                  {
                    key: 'A',
                    iconProps: { iconName: 'FitWidth' },
                    text: 'Nouveau template de produit',
                    checked: this.state.optionSelected === 'A'
                  },
                  {
                    key: 'B',
                    iconProps: { iconName: 'SearchAndApps' },
                    text: 'depuis Template',
                    checked: this.state.optionSelected === 'B'
                  }
                ]}
                onChange={this._onChange}
                required={true}
                disabled={this.state.isLoading}
              />
              

              {this.state.optionSelected === 'A' && (
                
                <div>
                  <TextField label="Nom de Template " value={this.state.newPageName} onChanged={this._onChangeNewPageName} placeholder="Insérer nom de  Template"
                readOnly={this.state.isLoading}
              />
                  
                </div>
              )}
              {this.state.optionSelected === 'B' && (
                
                <div>
                  <TextField label="Nom de Produit " value={this.state.newPageName} onChanged={this._onChangeNewPageName} placeholder="Insérer nom de  Produit"
                readOnly={this.state.isLoading}
              />
                  {console.log(this.state.allItems)}
                  <div>
                    {' '}
                    <Dropdown
                      label="Page Templates"
                      selectedKey={this.state.selectedItem ? this.state.selectedItem.key : undefined}
                      onChanged={this.changeState}
                      options={this.state.allItems}
                      disabled={this.state.isLoading}
                    />

                    {' '}
                  </div>
                </div>
              )}
              {!this.state.isLoading && (<div>
                <DialogFooter>
                  <PrimaryButton onClick={this._executeAction} text="Go!" />
                  <DefaultButton onClick={this._closeDialog} text="Cancel" />
                </DialogFooter>
              </div>)}
            </div>

          )}
          <Spinner size={SpinnerSize.large} label="Please, wait.." ariaLive="assertive" hidden={!this.state.isLoading} />
          {this.state.creationDone && (
            <div>
              <Label>Done!</Label>
              <DialogFooter>
                <PrimaryButton onClick={this._closeDialog} text="Close" />
              </DialogFooter>
            </div>
          )}
        </Dialog>
      </div>
    );
  }

  public changeState = (evt: any): void => {
    this.setState({ selectedItem: evt });
  }


  private _executeAction = (): void => {
    this.setState({ isLoading: true });
    let selected = this.state.selectedItem ? this.state.selectedItem.key : "";
    var resu: Promise<string> = ProductModelHelper.createCustomPage(this.state.newPageName, this.state.optionSelected, selected);
    resu.then(ss => {
      console.log(ss);
      this.setState({ isLoading: false, creationDone: true });
    });
  }

  private _onChangeNewPageName = (tmpPageName: any): void => {
    this.setState({ newPageName: tmpPageName });
  }



  private _onChange = (ev: React.FormEvent<HTMLInputElement>, option: any): void => {
    this.setState({ optionSelected: option.key });
  }

  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  }

  private _closeDialog = (): void => {
    this.props.close();
    this.setState({ hideDialog: true });

  }

} 