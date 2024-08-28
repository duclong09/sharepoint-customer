import { Log } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
//import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'CustomStyleApplicationCustomizerStrings';

const LOG_SOURCE: string = 'CustomStyleApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomStyleApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CustomStyleApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomStyleApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    SPComponentLoader.loadCss('https://shinegroupvn.sharepoint.com/sites/PORTAL/SiteAssets/customestyle.css');

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

   

    return Promise.resolve();
  }
}
