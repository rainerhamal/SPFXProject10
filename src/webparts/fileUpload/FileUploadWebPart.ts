import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './FileUploadWebPart.module.scss';
import * as strings from 'FileUploadWebPartStrings';

//! Lsn 4.7.6 Upload the file to the document library. Before we add the code to upload the file, we need to add a few references to object we'll use. 
import {
  ISPHttpClientOptions,
  SPHttpClient
} from '@microsoft/sp-http';

export interface IFileUploadWebPartProps {
  description: string;
}

export default class FileUploadWebPart extends BaseClientSideWebPart<IFileUploadWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  //! Lsn 4.7.2 With the web part created, update the user interface to include a control to select a file from the user's computer and a button to trigger the upload process.
  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.fileUpload} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
      </div>

      <div class="${styles.inputs}">
        <input class="${styles.fileUpload}-fileUpload" type="file"/><br/>
        <input class="${styles.fileUpload}-uploadButton" type="button" value="Upload"/>
      </div>

    </section>`;

    //? Lsn 4.7.2: TODO 1 These are used to get a reference to the two form elements you added to the web part:
    //* get reference to file control
    const inputFileElement = document.getElementsByClassName(`${styles.fileUpload}-fileUpload`)[0] as HTMLInputElement;
    //* wire up button control
    const uploadButton = document.getElementsByClassName(`${styles.fileUpload}-uploadButton`)[0] as HTMLButtonElement;

    //? Lsn 4.7.3: TODO 2 Now that you have a reference to the input control and button added to the web part, replace the // TODO 2 comment with the following:
    uploadButton.addEventListener('click', async () => {
      //* get filename
      const filePathParts = inputFileElement.value.split('\\');
      const fileName = filePathParts[filePathParts.length -1];

      //? Lsn 4.7.5 TODO 3 call this method _getFileBuffer() by adding the following line to our click event handler, immediately before the // TODO 3 comment:
      //* get file data: This call will take the first file selected by the user and pass it as a reference to the method we just added. The contents of the file are stored in the fileData member.
      const fileData = await this._getFileBuffer(inputFileElement.files[0]);

      //? Lsn 4.7.8 TODO 3 It calls the new _uploadFile() method and passes in the file's contents and the name of the file:
      //* upload file
      await this._uploadFile(fileData, fileName);
    });
  }

  //! Lsn 4.7.4 Read the contents of the selected file. This will take a file reference, read its contents into memory, and return it to the caller:
  private _getFileBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      let fileReader = new FileReader();
      // * write up error handler
      fileReader.onerror = (event: ProgressEvent<FileReader>) => {
        reject(event.target.error);
      };

      //* wire up when finished reading file
      fileReader.onloadend = (event: ProgressEvent<FileReader>) => {
        resolve(event.target.result as ArrayBuffer);
      };

      //* read file
      fileReader.readAsArrayBuffer(file);
    });
  }

  //! Lsn 4.7.7 Now, add the following method. This class will first create the full URL of the endpoint where you'll upload the file. Notice it's using the GetByTitle() method to reference the Documents library. It's also set to upload the file and overwrite an existing file with the same name. Next, after creating the request to send to the REST API endpoint, we're using the SpHttpClient object's post() method to upload the file to the SharePoint REST API. Once the file has been uploaded, an alert message notifies the user it worked. Otherwise it throws an exception.
  private async _uploadFile(fileData: ArrayBuffer, fileName: string): Promise<void> {
    //* create target endpoint for REST API HTTP POST
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Documents')/RootFolder/Files/add(overwrite=true,url='${fileName}')`;

    const options: ISPHttpClientOptions = {
      headers: { 'CONTENT-LENGTH': fileData.byteLength.toString() },
      body: fileData
    };

    //* upload file
    const response = await this.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, options);

    if (response.status === 200) {
      alert('File uploaded successfully');
    } else {
      throw new Error(`Error uploading file: ${response.statusText}`);
    }
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
