import * as React from 'react';
import styles from './WikiWebpartPnpjs.module.scss';
import { IWikiWebpartPnpjsProps } from './IWikiWebpartPnpjsProps';
import { sp } from '@pnp/sp'; 
import { 
  taxonomy, 
  ITermStore,
  ITerm,
  ITermData
} from '@pnp/sp-taxonomy'; 
import { Nav, INavLink, INavLinkGroup } from 'office-ui-fabric-react/lib/Nav';

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

  public async getSitePages2(store: ITermStore) {
    //récupérer toutes les pages
    var resp: Imms[] = await sp.web.lists.getByTitle("Site Pages").items.select("Title, FileRef, MMS").getAll(); 
    //récupérer toutes les données sur les termes
    const termSet = store.usingCaching().getTermSetById("452746d5-9636-4bc5-890f-473da11b1467"); 
    const select = ['IsRoot', 'Labels', 'TermsCount', 'Id', 'Name', 'Parent', 'LocalCustomProperties']; 
    const terms = await termSet.terms.select(...select).usingCaching().get();

    var respAl: Imms[] = []; 
    resp.forEach(element => {
      if(element["MMS"] != null) {
        respAl.push(element); 
      }
    });

    var bloglo: any[] = []; 

    var destination = respAl.map(function(item){
      
      if(item["MMS"] != null) {
        terms.forEach(element => {
          if((item["MMS"]["TermGuid"]).localeCompare(element.Id.substring(6, 42)) === 0) {
            const name = 'Parent';
            let one = {
              id: element.Id.substring(6, 42), 
              isRoot: element.IsRoot,  
              name: item["Title"], 
              parent: element[name] && element[name].Id ? element[name].Id.substring(6, 42): null, 
              path: item["FileRef"]
            }; 
            bloglo.push(one); 
          }
        });
      }
      return bloglo; 
    }).filter(function(item){return item}); 
    return destination[0]; 
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

        const sub: ISearchable = { termGuid: s.id, label: s.name, path: s.path, subTerms: [] };
        this.recursiveLink(sub, allTerms, link); 
        loopLink.links.push(link); 
      });
    }
  }

  public recursiveExpand(link: INavLink[]) {
    //console.log("je suis la")
    link.forEach(element => {
       
      if(0 == (element.url).localeCompare(window.location.href)) {
        //console.log("true"); 
        element.isExpanded = true; 
      } 
      else {
        //console.log("false"); 
        element.isExpanded = false; 
      }
    });
  }

  public isThere(loopLink: INavLink) {
    loopLink.links.forEach(element => {
      if((element.url).localeCompare(window.location.href)) {
        loopLink.isExpanded = true; 
        element.isExpanded = true; 
      }
    });
  }


  public componentDidMount() {
    const store = taxonomy.termStores.usingCaching().getById("a99d9ab5846d4dce891cd055c2b89690"); 
    this.getSitePages2(store).then((res: any[]) => {
      this.setState({
        terms: res
      })
    })
  }

  public render(): React.ReactElement<IWikiWebpartPnpjsProps> {

    const result = this.state.terms; 
    const rootTerms = result.filter((t) => t.isRoot === true);

    var link: INavLink[] = []; 
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
    })

    console.log("try"); 
    console.log(link); 
    this.recursiveExpand(link); 

    return ( 
      <div className={styles.wikiWebpartPnpjs}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Documentation Wiki</span>  
            </div>
          </div>
          <Nav
            //onLinkClick={_onLinkClick}
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
