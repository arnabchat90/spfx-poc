import * as React from 'react';
import {SPHttpClient,SPHttpClientConfigurations,HttpClient} from '@microsoft/sp-http';
import {ICustomSearchWebPartProps} from '../ICustomSearchWebPartProps';

export interface ISearchPanelProps extends ICustomSearchWebPartProps {
    httpClient : SPHttpClient;
    siteUrl : string;
}
