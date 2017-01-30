import * as React from 'react';
import {css, Button, ButtonType, TextField, autobind, Dropdown, IDropdownOption} from 'office-ui-fabric-react';
import {ICustomSearchWebPartProps} from '../ICustomSearchWebPartProps';
import {ISearchResultsProps} from './ISearchResultsProp';
import styles from './CustomSearch.module.scss';

export interface ISearchResultsState {
    //define states for the components
    status : string;
    formattedResults : ISearchResultRow[];
}

export interface ISearchColumn {
    Key : string;
    Value : string;
    ValueType : string;
}

export interface ISearchResultRow {
    DocTitle : string;
    DocUrl: string;
    DocAuthor : string;
    DocCaseRef : string;
    DocSubcaseRef : string;
    DocExtension : string;
    DocImage : string;
}

export default class SearchResults extends React.Component<ISearchResultsProps,ISearchResultsState> {
    private unFormattedRows = []
    constructor(props : ISearchResultsProps, state : ISearchResultsState) {
        super(props);
        //set your state with this.state
            this.state = {
                status : "",
                formattedResults : []
            }
        }
    public render() : JSX.Element {
        //getting results passed as property from Search results components
        this.unFormattedRows = this.props.results;
        this._formatResult(this.unFormattedRows, this);
        return(
        <div className='ms-Grid'>
            <div className='ms-Grid-row'>
             <div style={{height:'40px !important', lineHeight : '40px', textAlign: 'center'}} className='ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2 ms-bgColor-blueDark ms-fontColor-neutralLighterAlt ms-u-textAlignCenter ms-fontWeight-semibold'>Title</div>
             <div style={{height:'40px !important', lineHeight : '40px', textAlign: 'center'}} className='ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2 ms-bgColor-blueDark ms-fontColor-neutralLighterAlt ms-u-textAlignCenter ms-fontWeight-semibold'>Case Ref</div>
             <div style={{height:'40px !important', lineHeight : '40px', textAlign: 'center'}} className='ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2 ms-bgColor-blueDark ms-fontColor-neutralLighterAlt ms-u-textAlignCenter ms-fontWeight-semibold'>Subcase Ref</div>
             <div style={{height:'40px !important', lineHeight : '40px', textAlign: 'center'}} className='ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2 ms-bgColor-blueDark ms-fontColor-neutralLighterAlt ms-u-textAlignCenter ms-fontWeight-semibold'>Author</div>
             <div style={{height:'40px !important', lineHeight : '40px', textAlign: 'center'}} className='ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1 ms-bgColor-blueDark ms-fontColor-neutralLighterAlt ms-u-textAlignCenter ms-fontWeight-semibold'>Type</div>
             <div style={{height:'40px !important', lineHeight : '40px', textAlign: 'center'}} className='ms-Grid-col ms-u-sm3 ms-u-md3 ms-u-lg3 ms-bgColor-blueDark ms-fontColor-neutralLighterAlt ms-u-textAlignCenter ms-fontWeight-semibold'>Link</div>
            </div>
          <div className='ms-Grid-row'>{
              this.state.formattedResults.map(function(searchResult) {
                  return([
                       <div className='ms-Grid-row'>
                        <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2 ms-u-textAlignCenter ms-fontWeight-semibold">{searchResult.DocTitle}</div>
                        <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2 ms-u-textAlignCenter ms-fontWeight-semibold">{searchResult.DocCaseRef}</div>
                        <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2 ms-u-textAlignCenter ms-fontWeight-semibold">{searchResult.DocSubcaseRef}</div>
                        <div className="ms-Grid-col ms-u-sm2 ms-u-md2 ms-u-lg2 ms-u-textAlignCenter ms-fontWeight-semibold">{searchResult.DocAuthor}</div>
                        <div className="ms-Grid-col ms-u-sm1 ms-u-md1 ms-u-lg1 ms-u-textAlignCenter ms-fontWeight-semibold"> <div style={{display:'inline-block'}} className="ms-BrandIcon--Icon16" dangerouslySetInnerHTML={{__html: searchResult.DocImage}} /></div>
                        <div style={{display:'inline-block'}} className="ms-Grid-col ms-u-sm3 ms-u-md3 ms-u-lg3 ms-u-textAlignCenter ms-fontWeight-semibold"><a href={searchResult.DocUrl}>{searchResult.DocTitle}</a></div>
                       </div>
                  ]);
              })
          }
          </div>
          </div>
        );
    }

private _formatResult(unFormattedRows, self: SearchResults) {
    if(unFormattedRows != null) {
         self.state.formattedResults = [];
        unFormattedRows.map(function(row) {
                  var item : ISearchResultRow = {
                      DocTitle : "",
                      DocUrl : "",
                      DocCaseRef : "",
                      DocSubcaseRef : "",
                      DocAuthor : "",
                      DocExtension : "",
                      DocImage : ""
                  };
                  var cells : ISearchColumn[] = row.Cells;
                  for(var i=0; i<cells.length ; i++ ) {
                            if(cells[i].Key == "Title") {
                                item.DocTitle = cells[i].Value;
                            }
                            if(cells[i].Key == "Url") {
                                item.DocUrl = cells[i].Value;
                            }
                            if(cells[i].Key == "RefinableString90") {
                                item.DocCaseRef = cells[i].Value;
                            }
                            if(cells[i].Key == "RefinableString91") {
                                item.DocSubcaseRef = cells[i].Value;
                            }
                            if(cells[i].Key == "Author") {
                                item.DocAuthor = cells[i].Value;
                            }
                            if(cells[i].Key == "FileExtension") {
                                item.DocExtension = cells[i].Value;
                                switch(item.DocExtension) {
                                    case "docx" : 
                                        item.DocImage = '<div className="ms-BrandIcon--Icon16 ms-BrandIcon--Word"><img src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/word_16x1.svg" /></div>';
                                        break;
                                    case "xlsx" : 
                                        item.DocImage = '<div className="ms-BrandIcon--Icon16 ms-BrandIcon--Excel"><img src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_16x1.svg" /></div>';
                                        break;
                                    case "msg" : 
                                        item.DocImage = '<div className="ms-BrandIcon--Icon16 ms-BrandIcon--Outlook"><img src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/outlook_16x1.svg" /></div>';
                                        break;
                                    case "pptx" : 
                                        item.DocImage = '<div className="ms-BrandIcon--Icon16 ms-BrandIcon--Powerpoint"><img src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/powerpoint_16x1.svg" /></div>';
                                        break;
                                }
                            }
                            
                           
                  }
                   self.state.formattedResults.push(item);
              });
    }
    
}
  
}