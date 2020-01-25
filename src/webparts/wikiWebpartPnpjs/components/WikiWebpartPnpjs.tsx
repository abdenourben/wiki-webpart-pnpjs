import * as React from 'react';
import styles from './WikiWebpartPnpjs.module.scss';
import { IWikiWebpartPnpjsProps } from './IWikiWebpartPnpjsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp'; 
import { 
  taxonomy, 
  ITermStore,
  ITerms,
  ILabelMatchInfo,
  ITerm,
  ITermData } from '@pnp/sp-taxonomy'; 

export interface ISearchable {
  termGuid: string;
  label: string;
  path: any;
  subTerms: any;
}

export interface Imms {
  mms: {
    label: string;
    termGuid: string; 
    wssId: any; 
  }
}

export default class WikiWebpartPnpjs extends React.Component<IWikiWebpartPnpjsProps, {}> {

  public getSitePages() {
    sp.web.lists.getByTitle("Site Pages").items.select("Title, FileRef, MMS").getAll().then(async (resp: Imms[])=> {
      var termGuidTab: any[] = new Array(); 
      var termGuid: string; 
      var termTitle: string;
      var termName: string; 
      const store: ITermStore = await taxonomy.termStores.getById("a99d9ab5846d4dce891cd055c2b89690"); 
 

      resp.forEach(async element => {
        if(element["MMS"] != null ) {
          termGuidTab.push(element["MMS"]["TermGuid"]);
          termGuid = element["MMS"]["TermGuid"];
          termTitle = element["Title"]; 
          termName = element["FileRef"]; 

          var term: ITerm = store.getTermById(termGuid); 
          await term.setLocalCustomProperty("pageUrl", "https://m365x873105.sharepoint.com"+termName); 
          //const term2: ITerm & ITermData = await term.select("LocalCustomProperties").get();
          //console.log(term2["LocalCustomProperties"]["pageUrl"]); 
        }
      });
    });
    //sp.web.lists.getByTitle("Site Pages").items.getAll().then(console.log);
  }

  public async getPageUrlTerm(term: any) {
    const store: ITermStore = await taxonomy.termStores.getById("a99d9ab5846d4dce891cd055c2b89690"); 
    const termGuid = term.Id.substring(6, 42);
    var termOrigin: ITerm = store.getTermById(termGuid); 
    var termData: ITerm & ITermData = await termOrigin.select("LocalCustomProperties").get();
    if (termData["LocalCustomProperties"]["pageUrl"] != null)
      return termData["LocalCustomProperties"]["pageUrl"];
    else {
      return "?"; 
    }
  }

  public async getValuePromise(promise: Promise<any>) {
    return await promise; 
  }



  public async getSearchables() {
    const date = new Date(); 
    date.setDate(date.getDate() + 1); 

    const store = await taxonomy.termStores.usingCaching().getById("a99d9ab5846d4dce891cd055c2b89690"); 
    const termSet = await store.usingCaching().getTermSetById("452746d5-9636-4bc5-890f-473da11b1467"); 
    const select = ['IsRoot', 'Labels', 'TermsCount', 'Id', 'Name', 'Parent']; 
    const terms = await termSet.terms.select(...select).usingCaching().get();

    const allTerms: any[] = [
      ...terms.map(term => {
        //console.log("enter");
        const name = 'Parent';
        const promisePpageUrl = this.getPageUrlTerm(term);
    
        return {
          id: term.Id ? term.Id.substring(6, 42) : undefined, 
          isRoot: term.IsRoot, 
          name: term.Name, 
          parent: term[name] && term[name].Id ? term[name].Id.substring(6, 42): null, 
          path: promisePpageUrl
        };
        
      })
    ];    

    const searchables: ISearchable[] = []; 
    const rootTerms = allTerms.filter((t) => t.isRoot === true);
      const childTerms = allTerms
        .filter((t) => t.isRoot === false && t.parent)
        .sort((a,b) => a.parent - b.parent); 
  
      rootTerms.forEach((t) => {
        let term: ISearchable = { termGuid: t.id, label: t.name, path: t.path, subTerms: [] }; 
        term = this.recursiveAdd(term, childTerms); 
        term.path = t.path; 
        searchables.push(term);  
    });

    //this.searchables = searchables; 
    //console.log(searchables); 
    return(searchables) ; 
  }

  private recursiveAdd(currentTerm: any, allTerms: any): any {

    const subs = allTerms.filter((t) => {
      return t.parent != null && currentTerm.termGuid === t.parent;
    });

    if (subs != null && subs.length > 0) {
      currentTerm.subTerms = []; 
      subs.forEach((s) => {
        const sub: ISearchable = { termGuid: s.id, label: s.name, path: s.path, subTerms: [] };
        this.recursiveAdd(sub, allTerms); 
        //sub.path.pop(); 
        currentTerm.subTerms.push(sub);
      });
    }

    return currentTerm;
  }


  public componentDidMount() {
    this.getSitePages();  
    this.getSearchables().then(function(result){
      console.log(result); 
    }) 
  }

  public render(): React.ReactElement<IWikiWebpartPnpjsProps> {
    return (
      <div className={ styles.wikiWebpartPnpjs }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
