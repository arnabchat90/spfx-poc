import * as React from 'react';
import { css, Button, ButtonType, TextField, autobind, Dropdown, IDropdownOption } from 'office-ui-fabric-react';
import { ICustomSearchWebPartProps } from '../ICustomSearchWebPartProps';
import { ISearchPanelProps } from './ISearchPanelProps';
import styles from './CustomSearch.module.scss';
import { SPHttpClient, SPHttpClientConfigurations, HttpClient, SPHttpClientResponse, SPHttpClientConfiguration, IHttpClientConfiguration, HttpClientConfiguration, HttpClientConfigurations, ISPHttpClientConfiguration, ODataVersion } from '@microsoft/sp-http'
import SearchResults from "./SearchResults";
import { ISearchResultsProps } from './ISearchResultsProp';

export interface ISearchPanelState {
    //define states for the components
    status: string;
    items: any;
    rowCount: string;

}
const spSearchConfig: ISPHttpClientConfiguration = {
    defaultODataVersion: ODataVersion.v3
};
const clientConfigODataV3: SPHttpClientConfiguration = SPHttpClient.configurations.v1.overrideWith(spSearchConfig);
export default class SearchPanel extends React.Component<ISearchPanelProps, ISearchPanelState> {
    private varCaseRef: string;
    private varSubcaseRef: string;
    private varQuery; string;
    private selectedFileType: string;
    constructor(props: ISearchPanelProps, state: ISearchPanelState) {
        super(props);
        this.varCaseRef = null;
        this.varSubcaseRef = null;
        this.varQuery = null;
        this.selectedFileType = "docx";
        //set your state with this.state
        this.state = {
            status: "",
            items: null,
            rowCount: ""
        }

    }
    public render(): JSX.Element {
        return (
            <div>
                <div className='ms-Grid'>
                    <div className='ms-Grid-row'>
                        <TextField style={{ height: '50px!important' }} label='Search This Site' placeholder='Perform a Search...' ariaLabel='Please enter text here' onChanged={this._onChangedTxtSearch} />
                    </div>
                    <div className='ms-Grid-row'>
                        <div className="ms-Grid-col ms-u-sm4 ms-u-md4 ms-u-lg4">
                            <TextField id='txtCaseRef' label='Case Reference' placeholder='Case Refs starts with C00XX' onChanged={this._onChangedTxtCaseRef} />
                        </div>
                        <div className="ms-Grid-col ms-u-sm4 ms-u-md4 ms-u-lg4">
                            <TextField id='txtSubcaseRef' label='Subcase Reference' placeholder='Subcase Refs starts with SC00XXX' onChanged={this._onChangedTxtSubcaseRef} />
                        </div>
                        <div className="ms-Grid-col ms-u-sm4 ms-u-md4 ms-u-lg4">
                            <Dropdown
                                label='Document Type:' selectedKey='docx'
                                options={
                                    [
                                        { key: 'docx', text: 'Documents' },
                                        { key: 'pptx', text: 'PowerPoint Presentations' },
                                        { key: 'xlsx', text: 'Excels' },
                                        { key: 'msg', text: 'Emails' },
                                        { key: 'pdf', text: 'PDF' }
                                    ]
                                }
                                onChanged={(item) => {
                                    this.selectedFileType = item.key as string;
                                }
                                }
                                />
                        </div>
                    </div>
                    <div style={{ marginTop: '20px' }} className='ms-Grid-row'>
                        <div className="ms-Grid-col ms-u-sm6 ms-u-smPush4">
                            <Button style={{ width: '240px' }} id='btnSearch' buttonType={ButtonType.primary} onClick={this.doSearch}>Search SharePoint</Button>
                        </div>
                    </div>

                    <div style={{ marginTop: '20px' }} className='ms-Grid-row'>
                        <div className="ms-Grid-col ms-u-sm12">
                            {this.state.rowCount}
                        </div>

                    </div>

                </div>
                <div>
                    <SearchResults description="Search results" results={this.state.items} />
                </div>
            </div>
        );
    }

    @autobind
    private _onChangedTxtSearch(text) {
        this.varQuery = text;
        console.log(this.varQuery);
    }
    @autobind
    private _onChangedTxtCaseRef(text) {
        this.varCaseRef = text;
        console.log(this.varCaseRef);
    }
    @autobind
    private _onChangedTxtSubcaseRef(text) {
        this.varSubcaseRef = text;
        console.log(this.varSubcaseRef);
    }
    @autobind
    private doSearch(): Promise<void> {
        this.setState({
            status: 'Loading latest items...',
            items: null,
            rowCount: "Loading Items"
        });
        console.log(this.state.status);
        //Logic to create URL
        // var url="https://nvsdev.sharepoint.com/sites/spfx-dev/_api/web/lists";
        var url = `${this.props.siteUrl}/_api/search/query?querytext='${this.varQuery}*'&selectproperties='Title,Author,Url,RefinableString90,RefinableString91,FileExtension'`;
        var fileTypeRefiners = "";
        var fileTypeRefinementFilters = "";
        var caseRefRefiners = "";
        var caseRefRefinementFilters = "";
        var subcaseRefRefiners = "";
        var subcaseRefRefinementFilters = "";
        if(this.selectedFileType != null && this.selectedFileType != "") {
            fileTypeRefiners += "filetype"
            fileTypeRefinementFilters += `filetype:equals(\"${this.selectedFileType}\")`;
        }
        if(this.varCaseRef != null && this.varCaseRef != "") {
           caseRefRefiners += "RefinableString90";
           caseRefRefinementFilters += `RefinableString90:equals(\"${this.varCaseRef}\")`;
        }
        if(this.varSubcaseRef != null && this.varSubcaseRef != "") {
           subcaseRefRefiners += "RefinableString91";
           subcaseRefRefinementFilters += `RefinableString91:equals(\"${this.varSubcaseRef}\")`;
        }
        url += `&refiners='${fileTypeRefiners},${caseRefRefiners},${subcaseRefRefiners}`;
        url = url.replace(/(^[,\s]+)|([,\s]+$)/g, '');
        url = url.replace(/[, ]+/g, ",").trim();
        url += "'";
        if(caseRefRefinementFilters != null && caseRefRefinementFilters != "" && subcaseRefRefinementFilters == "") {
            url += `&refinementfilters='and(${fileTypeRefinementFilters},${caseRefRefinementFilters})'`;   
        }
        else if(subcaseRefRefinementFilters != null && subcaseRefRefinementFilters != "") {
            url += `&refinementfilters='and(${fileTypeRefinementFilters},${caseRefRefinementFilters},${subcaseRefRefinementFilters})'`;
            url = url.replace(/(^[,\s]+)|([,\s]+$)/g, '');
            url = url.replace(/[, ]+/g, ",").trim();   
        }
        else {
            url += `&refinementfilters='${fileTypeRefinementFilters}'`;
        }
        // if(this.selectedFileType != null) {
        //     url = url + `&refiners='filetype'&refinementfilters='filetype:equals(\"${this.selectedFileType.key}\")'`;
        // }
        // if(this.varCaseRef != null) {
        //     url = url + ``
        // }
        return this.props.httpClient.get(url, clientConfigODataV3, {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': ''
            }
        })
            .then((response: SPHttpClientResponse): Promise<{ value: any }> => {
                console.log(this.state.status);
                return response.json();

            })
            .then((responseJSON: any): void => {
                this.setState({
                    status: `Successfully loaded items`,
                    rowCount: "Total Rows - " + responseJSON.PrimaryQueryResult.RelevantResults.Table.Rows.length,
                    items: null
                });
                if (responseJSON.PrimaryQueryResult.RelevantResults.RowCount != null || this.state.items.PrimaryQueryResult.RelevantResults.RowCount != undefined || this.state.items.PrimaryQueryResult.RelevantResults.RowCount > 0) {
                    //pass the rows to SearchResults
                    this.setState({
                        status: `Successfully loaded items`,
                        rowCount: "Total Rows - " + responseJSON.PrimaryQueryResult.RelevantResults.Table.Rows.length,
                        items: responseJSON.PrimaryQueryResult.RelevantResults.Table.Rows
                    });
                }
                console.log(this.state.status);
            }, (error: any): void => {
                this.setState({
                    status: 'Loading all items failed with error: ' + error,
                    items: null,
                    rowCount: "Error Occured"
                });
                console.log(this.state.status);
            });
    }
}
