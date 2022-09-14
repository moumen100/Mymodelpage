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
  /**
   * Constructeur pour initialiser  state de component React
   * @param props Objet des Props de component React
   */
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
    /**
     * récupérer les types de modèles qui existe dans la liste sharepoint  "PageModelType"
     * utilisons pnpJs afin de consummer les ressources SharePoint
     */
    sp.web.lists.getByTitle("PageModelType").items.select("Id,Title,Description").getAll().then((items: IModelsTypes[]) => {
      var tmpItems: IModelsTypes[] = new Array();
      //DropDown initialization
      items.forEach(element => {
        var item = { Id: element["Id"], Title: element["Title"], Description: element["Description"] };
        tmpItems.push(item);  // assigner la liste des modèles sur la liste déroullante  
      });
      this.setState({ ModelsTypes: tmpItems }); // mettre à jour la liste de smodèles en changeant le state 
      console.log("Checking values ----");
      console.log( this.state.ModelsTypes)
    });
  }
  /**
   * 
   * @param ModelType récupérer la liste des pages en fonction de type de template choisis
   * @returns retournner la liste des modèles des pages
   */
  public async GetTemplates(ModelType:string) : Promise<IpageModel[]> {
    console.log("GetTemplates Start with param : "+ ModelType)
    var tmpItems: IpageModel[] = new Array();
    //récupérer la liste des Pages depuis la library des pages sur SharePoint  "Site Pages"  en filtrant par le Titre de type de modèle
    sp.web.lists.getByTitle("Site Pages").items.expand("ModelType").select("Title,FileRef,ModelType/Title").filter(`ModelTypeId eq '${ModelType}'`).getAll().then((items: any[]) => {
      //DropDown initialization
      console.log("Requests Values " + items)
      //remplir la liste déroullante contient la liste des templates
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
                            {/* choix 1 pour créer un nouveau modèle  
                                choix 2 pour créer a partie d'une template
                            */}
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
                onChange={this._onChange} //en change le choix 
                required={true}
                disabled={this.state.isLoading}
              />
              
                {/* contenu de choix 1 : créer un nouveau modèle */}
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
              {/* contenu de choix 2 : créer à partie d'un modèle  */}
              {this.state.optionSelected === 'B' && (
                  // titre de la page
                <div>
                  <TextField label="Intitulé de la page" value={this.state.newPageName} onChanged={this._onChangeNewPageName} placeholder="Insérer nom de Produit"
                readOnly={this.state.isLoading}
              />
              {/* les catégories de modèle */}
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
                    {/* La liste des template */}
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
              {/* footer de dialogue contient les Buttons d'actions Comfirer la création ou Annuller */}
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
  /**
   * méthode s'execute lors de cliquer sur créer sur la dialogue 
   */
  private _executeAction = (): void => {
    this.setState({ isLoading: true });
    let selected = this.state.selectedModel ? this.state.selectedModel.key : "";
    let ModelTypeSelected ; if ( this.state.selectedModelsType)  { ModelTypeSelected = this.state.selectedModelsType.key }else {ModelTypeSelected =  ""};
    const ModelHelper = new ProductModelHelper()
    // / Faire appel à la méthode createCustomPage permet de créer un modèle ou apartir d'un modèle
    // Prend en paramètre titre, 
    // option selectioné : Soit créer un modèle soit apartir d'un modèle
    //type de modèle selectionné 
    var resu: Promise<string> = ModelHelper.createCustomPage(this.state.newPageName, this.state.optionSelected, selected, ModelTypeSelected);
    // roulle d'attente
    resu.then(ss => {
      console.log(ss);
      this.setState({ isLoading: false, creationDone: true });
    });
  }

  /**
   * évenement déclanche à chaque changement de type d'action 
   * @param evt  Objet  évènement de JS
   */
  public changeState = (evt: any): void => {
    this.setState({ selectedModel: evt })
  }
  /**
   * évenement déclanche à chaque changement de type de modèle utilisé
   * 
   * @param evt Objet  évènement de JS
   */
  public  changeModelTypeState = (evt: any): void => {
    console.log("Model type change event triigered")
     this.GetTemplates(evt.key).then( ListTemplate => {
      this.setState({ selectedModelsType:evt,allIModels:ListTemplate }); // changement de laliste des modèles en fonction de catégory choisis
     })
  }
  /**
   * évenement déclanche à chaque changement de titre de la page
   * @param tmpPageName nom de la page
   */
  private _onChangeNewPageName = (tmpPageName: any): void => {
    this.setState({ newPageName: tmpPageName });
  }

  /**
   * évènement en change type d'action 
   * @param ev 
   * @param option 
   */
  private _onChange = (ev: React.FormEvent<HTMLInputElement>, option: any): void => {
    this.setState({ optionSelected: option.key });
  }
  /**
   * 
   */
  private _closeDialog = (): void => {
    this.props.close();
    this.setState({ hideDialog: true });
    window.location.reload();


  }

} 