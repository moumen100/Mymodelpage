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
import IpageModel from '../Models/IPageModels';
import IModelsTypes from '../Models/IModelsTypes';

export interface BBS_CPA_IProductModelDialogContentState {
  hideDialog: boolean;
  optionSelected: string;
  newPageName: string;
  selectedModel?: { key: string | number | undefined };
  selectedModels: string[];
  allIModels: IpageModel[];
  isLoading: boolean;
  creationDone: boolean;
  ModelsTypes : IModelsTypes[]
  selectedModelsType : { key: string | number | undefined };
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
      selectedModel: null,
      selectedModels: null,
      allIModels: [],
      isLoading: false,
      creationDone: false,
      ModelsTypes:[],
      selectedModelsType : null
    };
    
    sp.web.lists.getByTitle("PageModelType").items.select("Id,Title,Description").getAll().then((items: IModelsTypes[]) => {
      var tmpItems: IModelsTypes[] = new Array();
      //DropDown initialization
      items.forEach(element => {
        var item = { Id: element["Id"], Title: element["Title"], Description: element["Description"] };
        tmpItems.push(item);
      });
      this.setState({ ModelsTypes: tmpItems });
      console.log("Checking values ----");
      console.log( this.state.ModelsTypes)
    });
  }
  
  public async GetTemplates(ModelType:string) : Promise<IpageModel[]> {
    console.log("GetTemplates Start with param : "+ ModelType)
    var tmpItems: IpageModel[] = new Array();
    // .filter(`TypeTitle eq '${ModelType}'`)
    sp.web.lists.getByTitle("Site Pages").items.expand("ModelType").select("Title,FileRef,ModelType/Title").filter(`ModelTypeId eq '${ModelType}'`).getAll().then((items: any[]) => {
      //DropDown initialization
      console.log("Requests Values " + items)
      items.forEach(element => {
        var item = { key: element["FileRef"], text: element["Title"] };
        tmpItems.push(item);
      });
      console.log("Checking values ----");
        console.log(tmpItems);
      
    });
    return tmpItems
  }

  public render() {

    return (
      <div>
        <Dialog
          hidden={this.state.hideDialog}

          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: "Créer une page à partir d'un modèle"
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
                    text: 'Nouveau modèle',
                    checked: this.state.optionSelected === 'A'
                  },
                  {
                    key: 'B',
                    iconProps: { iconName: 'SearchAndApps' },
                    text: 'Depuis un modèle ',
                    checked: this.state.optionSelected === 'B'
                  }
                ]}
                onChange={this._onChange}
                required={true}
                disabled={this.state.isLoading}
              />
              

              {this.state.optionSelected === 'A' && (
                
                <div>
                  <TextField label="Intitulé de modèle " value={this.state.newPageName} onChanged={this._onChangeNewPageName} placeholder="Insérer nom de  Template"
                readOnly={this.state.isLoading}
              />
                  <Dropdown
                      label="Catégorie du modèle"
                      selectedKey={this.state.selectedModelsType ? this.state.selectedModelsType.key : undefined}
                      onChanged= { this.changeModelTypeState}
                      options={this.state.ModelsTypes.map(m=> {return {key:m.Id, text:m.Title}})}
                      disabled={this.state.isLoading}
                    />
                </div>

              )}
              {this.state.optionSelected === 'B' && (
                
                <div>
                  <TextField label="Intitulé de la page" value={this.state.newPageName} onChanged={this._onChangeNewPageName} placeholder="Insérer nom de Produit"
                readOnly={this.state.isLoading}
              />
              <Dropdown
                      label="Catégorie du modèle "
                      selectedKey={this.state.selectedModelsType ? this.state.selectedModelsType.key : undefined}
                      onChanged= { this.changeModelTypeState}
                      options={this.state.ModelsTypes.map(m=> {return {key:m.Id, text:m.Title}})}
                      disabled={this.state.isLoading}
                    />
                  {console.log(this.state.allIModels)}
                  <div>
                    {' '}
                    <Dropdown
                      label="Modèle"
                      selectedKey={this.state.selectedModel ? this.state.selectedModel.key : undefined}
                      onChanged={this.changeState}
                      options={this.state.allIModels}
                      disabled={this.state.isLoading}
                    />

                    {' '}
                  </div>
                </div>
              )}
              {!this.state.isLoading && (<div>
                <DialogFooter>
                  <PrimaryButton onClick={this._executeAction} text="Créer" />
                  <DefaultButton onClick={this._closeDialog} text="Annuler" />
                </DialogFooter>
              </div>)}
            </div>

          )}
          <Spinner size={SpinnerSize.large} label="Please, wait.." ariaLive="assertive" hidden={!this.state.isLoading} />
          {this.state.creationDone && (
            <div>
              <Label>Votre modèle a été bien créé</Label>
              <DialogFooter>
                <PrimaryButton onClick={this._closeDialog} text="Fermer" />
              </DialogFooter>
            </div>
          )}
        </Dialog>
      </div>
    );
  }

  public changeState = (evt: any): void => {
    this.setState({ selectedModel: evt });
  }
  public  changeModelTypeState = (evt: any): void => {
    console.log("Model type change event triigered")
     this.GetTemplates(evt.key).then( ListTemplate => {
      this.setState({ selectedModelsType:evt,allIModels:ListTemplate  });
     })
  }

  private _executeAction = (): void => {
    this.setState({ isLoading: true });
    let selected = this.state.selectedModel ? this.state.selectedModel.key : "";
    let ModelTypeSelected ; if ( this.state.selectedModelsType)  { ModelTypeSelected = this.state.selectedModelsType.key }else {ModelTypeSelected =  ""};
    const ModelHelper = new ProductModelHelper()
    var resu: Promise<string> = ModelHelper.createCustomPage(this.state.newPageName, this.state.optionSelected, selected, ModelTypeSelected);
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
    window.location.reload();


  }

} 