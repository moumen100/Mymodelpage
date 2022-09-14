import {
    ClientSideText, 
    ClientSideWebpart,
    sp,
    ClientSidePage,
    Field,
    Fields
} from "@pnp/sp";

export declare namespace MyClientSideWebpartPropertyTypes {
    /**
     * Properties for People (component id: 7f718435-ee4d-431c-bdbf-9c4ff326f46e)
     */
    interface People {
        layout: "1" | "2";
        persons?: any[];
    }
}

export class ProductModelHelper {
   /**
    * 
    * @param pagename Nom de la page insérer
    * @param pageType  le Type d'action à faire nouveau ou apartir d'un modèle
    * @param templatePageUrl  // url de la page comme étant modèle
    * @param ModelTypeId // Id de catégory de modèle
    * @returns 
    */
    public  async createCustomPage(pagename: string, pageType: string, templatePageUrl: any, ModelTypeId:any): Promise<string> {
        switch (pageType) {
            case 'A': // cas 1 : créer une nouveau modèle
                //créer une page avec pnpJs 
                const page = await sp.web.addClientSidePage(pagename + ".aspx");
                console.log("pagetype" + pageType);
                //récupérer le webParts
                const partDefs = await sp.web.getClientSideWebParts();
                console.log("case a");
                const section = page.addSection(); // Ajouter une section sur la page 
                console.log("section added");

                const column1 = section.addColumn(4);

                // trouver la webPart par Id
                const partDef = partDefs.filter(c => c.Id === "7f718435-ee4d-431c-bdbf-9c4ff326f46e");
                
                // tester si la webpart trouvé
                if (partDef.length < 1) {
                    // retourner une erreur
                    console.log('ops');
                    throw new Error("Could not find the web part");
                }
                // créer une instance de la WebPart ( User)
                const part = ClientSideWebpart.fromComponentDef(partDef[0]);

                part.setProperties<MyClientSideWebpartPropertyTypes.People>({
                    layout: "2",
                    persons: [
                        {
                            "id": "i:0#.f|membership|jsmith@federicoporceddu.onmicrosoft.com",
                            "upn": "jsmith@federicoporceddu.onmicroosft.com",
                            "role": "",
                            "department": "",
                            "phone": "",
                            "sip": ""
                        }
                    ]
                });
                // ajouter la WebPart 
                column1.addControl(part);

                const column2 = section.addColumn(8);
                // ajouter un text à la 2ème collone
                column2.addControl(new ClientSideText("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."));

                page.disableComments();
                await page.save();
                await this.GetUpdatePage("ModelTypeId", ModelTypeId,pagename)
               
                console.log("case saved");
                break;
            case 'B':
                // cas créer apartir d'un modèle
                console.log("case b - start");
                // récupérer la page  ( template ) selectinné 
                const templatePage = await ClientSidePage.fromFile(sp.web.getFileByServerRelativeUrl(templatePageUrl));
                //créer une copie de la page
                await templatePage.copyPage(sp.web, pagename + ".aspx", pagename, false);
                console.log("case b - end");

                // Mettre à jour les collones 
                await this.GetUpdatePage("PageTypeId", ModelTypeId,pagename)
                break;
            default:
                break;
        }

        return await "done";
    }



    /**
     * Cette méthode permet de modifier les champs sur les éléments d'une library sharepoint
     * @param fieldname le chmaps à mettre à jour
     * @param newvalue nouvelle valeur
     * @param pageTitle Titre de la page à modifier
     */
    public  async GetUpdatePage(fieldname:string, newvalue:any,pageTitle:string) : Promise<void> {
        console.log("fields: ")
        console.log(fieldname)

        const pageObj = new Object()

         pageObj[fieldname] = newvalue 
            console.log("pageObje:")
            console.log(pageObj)

            const Pages = await sp.web.lists.getByTitle("Site Pages") // récupérer la page avec pnpJs
            
            const MyPAge: any[] = await Pages.items.top(1).filter("Title eq '"+pageTitle+"'").get()     
            console.log("MyPage: ")
            console.log(MyPAge)
        if (MyPAge.length>0) {
                                await Pages.items.getById(MyPAge[0].Id).update({...pageObj}) // mettre à jour le champs
                             }
    }
    
    /**
     * récupérer le'Url d'une page sharepoint 
     * @param pageTitle titre de la page
     * @returns retourner l'Url
     */
    public  async GetPageUrl(pageTitle:string) : Promise<string> {
        console.log("start geturl function")
        const Pages = await sp.web.lists.getByTitle("Site Pages")
        const MyPAge: any[] = await Pages.items.top(1).filter("Title eq '"+pageTitle+"'").select("Title,FileRef").get()     
        console.log(MyPAge)
    if (MyPAge.length>0) {
        console.log(MyPAge[0].FileRef)
        console.log(sp.web.toUrl())
                                return sp.web.getFileByServerRelativeUrl(MyPAge[0].FileRef).toUrl()                        
                            }
        
}
}

