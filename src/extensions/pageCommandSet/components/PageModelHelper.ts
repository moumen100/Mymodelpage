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
    private MyWeb;
    constructor(Web) {
        this.MyWeb = Web
    }

    // public static async  getInfos(pagename: string): Promise<string> {

    //     var resultData: any = await  .lists.getByTitle("Site Pages")
    //         .items.getById(15)
    //         .select("Title")
    //         .get();

    //     return await resultData.Title;
    // }

    public  async createCustomPage(pagename: string, pageType: string, templatePageUrl: any, ModelTypeId:any): Promise<string> {
        switch (pageType) {
            case 'A':
                const page = await this.MyWeb.addClientSidePage(pagename + ".aspx");
                console.log("pagetype" + pageType);
                
                const partDefs = await this.MyWeb.getClientSideWebParts();
                console.log("case a");
                const section = page.addSection();
                console.log("section added");

                const column1 = section.addColumn(4);

                // find the definition we want, here by id
                const partDef = partDefs.filter(c => c.Id === "7f718435-ee4d-431c-bdbf-9c4ff326f46e");

                // optionally ensure you found the def
                if (partDef.length < 1) {
                    // we didn't find it so we throw an error
                    console.log('ops');
                    throw new Error("Could not find the web part");
                }
                // create a ClientWebPart instance from the definition
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
                // add a text control to the second new column
                column1.addControl(part);

                const column2 = section.addColumn(8);
                //// add a text control to the second new column
                column2.addControl(new ClientSideText("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."));

                page.disableComments();
                await page.save();
                await this.GetUpdatePage("ModelTypeId", ModelTypeId,pagename)
                // const e= await this.GetPageUrl(pagename)
                // console.log("PageUrl: ")
                // console.log(e)
                console.log("case saved");
                break;
            case 'B':
                console.log("case b - start");
                const templatePage = await ClientSidePage.fromFile(this.MyWeb.getFileByServerRelativeUrl(templatePageUrl));
                await templatePage.copyPage(this.MyWeb, pagename + ".aspx", pagename, false);
                console.log("case b - end");

                // window.location.href= await this.GetPageUrl(pagename)
                await this.GetUpdatePage("PageTypeId", ModelTypeId,pagename)
                break;
            default:
                break;
        }

        return await "done";
    }




    public  async GetUpdatePage(fieldname:string, newvalue:any,pageTitle:string) : Promise<void> {
        console.log("fields: ")
        console.log(fieldname)

        const pageObj = new Object()

         pageObj[fieldname] = newvalue 
            console.log("pageObje:")
            console.log(pageObj)

            const Pages = await this.MyWeb.lists.getByTitle("Site Pages")

            const MyPAge: any[] = await Pages.items.top(1).filter("Title eq '"+pageTitle+"'").get()     
            console.log("MyPage: ")
            console.log(MyPAge)
        if (MyPAge.length>0) {
                                await Pages.items.getById(MyPAge[0].Id).update({...pageObj})
                             }
            
    }
    

    public  async GetPageUrl(pageTitle:string) : Promise<string> {
        console.log("start geturl function")
        const Pages = await this.MyWeb.lists.getByTitle("Site Pages")
        const MyPAge: any[] = await Pages.items.top(1).filter("Title eq '"+pageTitle+"'").select("Title,FileRef").get()     
        console.log(MyPAge)
    if (MyPAge.length>0) {
        console.log(MyPAge[0].FileRef)
        console.log(this.MyWeb.toUrl())
                                return this.MyWeb.getFileByServerRelativeUrl(MyPAge[0].FileRef).toUrl()                        }
        
}
}

