import * as React from 'react';
import styles from './WikiWebpartPnpjs.module.scss';
import { IWikiWebpartPnpjsProps } from './IWikiWebpartPnpjsProps';
import { sp } from '@pnp/sp'; 
import { 
  taxonomy, 
  ITermStore,
  ITerm
} from '@pnp/sp-taxonomy'; 
import { Nav, INavLink } from 'office-ui-fabric-react/lib/Nav';

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

export interface ITermsState {
  terms: ItermX[]
}

export interface ItermX {
  id: any; 
  isRoot: any; 
  name: any;
  parent: any;
  path: any
}

export interface Group {
  name: any;
  links: Link[]; 
}

export interface Link {
  name: any;
  url: any;
  target: any;
  expandAriaLabel: string,
  collapseAriaLabel: string,
  links: Link[]; 
}

export default class WikiWebpartPnpjs extends React.Component<IWikiWebpartPnpjsProps, ITermsState > {
  constructor(props) {
    super(props);
    this.state = {
      terms: []
    }
  }


  public getSitePages(store: ITermStore) {
    sp.web.lists.getByTitle("Site Pages").items.select("Title, FileRef, MMS").getAll().then((resp: Imms[])=> {
      var termGuidTab: any[] = new Array(); 
      var termGuid: string, termTitle: string, termName: string;    
      resp.forEach(async element => {
        if(element["MMS"] != null ) {
          termGuidTab.push(element["MMS"]["TermGuid"]);
          termGuid = element["MMS"]["TermGuid"];
          termTitle = element["Title"]; 
          termName = element["FileRef"]; 
          var term: ITerm = store.getTermById(termGuid); 
          term.setLocalCustomProperty("titlePage", termTitle);
          term.setLocalCustomProperty("pageUrl", "https://m365x873105.sharepoint.com"+termName); 
        }
      });
    });
  }


  public async getTerms() {
    const date = new Date(); 
    date.setDate(date.getDate() + 1); 

    const store = await taxonomy.termStores.usingCaching().getById("a99d9ab5846d4dce891cd055c2b89690"); 
    this.getSitePages(store); 
    const termSet = await store.usingCaching().getTermSetById("452746d5-9636-4bc5-890f-473da11b1467"); 
    const select = ['IsRoot', 'Labels', 'TermsCount', 'Id', 'Name', 'Parent', 'LocalCustomProperties']; 
    const terms = await termSet.terms.select(...select).usingCaching().get();
    const allTerms: any[] = [
      ...terms.map(term => {
        const name = 'Parent';
        return {
            id: term.Id ? term.Id.substring(6, 42) : undefined, 
            isRoot: term.IsRoot, 
            name: term.LocalCustomProperties["titlePage"], 
            parent: term[name] && term[name].Id ? term[name].Id.substring(6, 42): null, 
            path: term.LocalCustomProperties["pageUrl"]
          };
      })
    ];
    return(allTerms); 
  } 


  public recursiveLink(currentTerm: any, allTerms: any, loopLink: INavLink) {
    const subs = allTerms.filter((t) => {
      return t.parent != null && currentTerm.termGuid === t.parent;
    });
    if (subs != null && subs.length > 0) {
      subs.forEach(s => {
          const link: INavLink = {
          name: s.name, 
          url: s.path, 
          target: s.path,
          links: []
        };

        link.isExpanded = true; 

        const sub: ISearchable = { termGuid: s.id, label: s.name, path: s.path, subTerms: [] };
        this.recursiveLink(sub, allTerms, link)
        loopLink.links.push(link); 
      });
    }
  }


  public componentDidMount() {
    const store = taxonomy.termStores.usingCaching().getById("a99d9ab5846d4dce891cd055c2b89690"); 
    this.getSitePages(store); 
    this.getTerms().then((res: any[]) => {
      this.setState({
        terms: res
      })
    }); 
  }


  public render(): React.ReactElement<IWikiWebpartPnpjsProps> {
   
    const result = this.state.terms; 
    var link: INavLink[] = []; 
    const rootTerms = result.filter((t) => t.isRoot === true);
    rootTerms.forEach((t) => {
      let term: ISearchable = { termGuid: t.id, label: t.name, path: t.path, subTerms: [] }; 
      let loopLink: INavLink = {
        name: t.name,
        url: t.path,
        target: t.path,
        links: [],
      };
      this.recursiveLink(term, result, loopLink); 
      term.path = t.path; 
      link.push(loopLink); 
  });

  link.forEach(element => {
    element.isExpanded = true; 
  });

  function _onLinkClick(ev: React.MouseEvent<HTMLElement>, item?: INavLink) {
    if (item && item.name === 'Version Mac') {
      alert('Version Mac clicked');
      item.isExpanded = false; 
    }
    console.log(item.isExpanded); 
    console.log(item.parentId); 
  }


    return ( 
      <div className={styles.wikiWebpartPnpjs}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Documentation Wiki</span>  
            </div>
          </div>
          <Nav
            onLinkClick={_onLinkClick}
            ariaLabel="Nav example with nested links"
            groups={[
              {
                links: link
              }
            ]}
          />
        </div>
      </div>
    );
  }
}
