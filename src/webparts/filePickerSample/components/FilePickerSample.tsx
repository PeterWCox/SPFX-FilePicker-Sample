import * as React from "react";
import { IFilePickerSampleWebPartProps } from "../FilePickerSampleWebPart";
import { SharePointDocumentLibraryRepository } from "../../../repositories/SharepointDocumentLibraryRepository";
import { IFileAddResult } from "@pnp/sp/files";
import { FilePicker, IFilePickerResult } from "@pnp/spfx-controls-react";
import {
  Link,
  MessageBarType,
  Text,
  Image,
} from "@fluentui/react";
import { MessageBar } from "office-ui-fabric-react";

export interface IFilePickerSampleProps {
  wpp: IFilePickerSampleWebPartProps;
  context: any;
}

export const FilePickerSample = (props: IFilePickerSampleProps) => {
  const [imageUrl, setImageUrl] = React.useState<string>("");
  const [isImagePickerVisible, setIsImagePickerVisible] =
    React.useState<boolean>(false);

  const [urls, setUrls] = React.useState<string[]>([]);
  const [isUrlPickerVisible, setIsUrlPickerVisible] =
    React.useState<boolean>(false);

  const [
    isImageUploadErrorMessageDisplayed,
    setIsImageUploadErrorMessageDisplayed,
  ] = React.useState<boolean>(false);

  //Event handler when we're uploading Img's (could also be used for other document types...)
  const onImagePickerSave = async (fpr: IFilePickerResult[]): Promise<void> => {
    setIsImageUploadErrorMessageDisplayed(false);

    //Only consider the first file picker result
    const filePickerResult = fpr[0];

    //If a stock image or an existing document library image was selected
    if (filePickerResult.fileAbsoluteUrl) {
      setImageUrl(filePickerResult.fileAbsoluteUrl);
    }

    //Otherwise its a locally uploaded file...
    const spDocumentLibraryRepository = new SharePointDocumentLibraryRepository(
      props.context
    );
    const response: IFileAddResult =
      await spDocumentLibraryRepository.uploadFileToSharePointDocumentLibrary(
        filePickerResult,
        props.wpp.imageStorageSharePointDocumentLibrary
      );

    //If successful (i.e. response !== undefined, package does not send a response if error!)
    if (response) {
      setImageUrl(document.location.origin + response.data.ServerRelativeUrl);
    } else {
      //Inform user that the image upload failed
      setIsImageUploadErrorMessageDisplayed(true);
    }
  };

  const onUrlPickerSave = async (fpr: IFilePickerResult[]): Promise<void> => {
    fpr.forEach(filePickerResult => {
      setUrls(urls => [...urls, filePickerResult.fileAbsoluteUrl]);
    });
  };

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        gap: "20px",
      }}
    >
      <div>
        <Text>Image Picker Demo</Text>

        <FilePicker
          // accepts={["",".gif", ".jpg", ".jpeg", ".bmp", ".dib" , ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"]}
          hideLinkUploadTab
          // hideLocalUploadTab
          hideOneDriveTab
          hideWebSearchTab
          disabled={
            !props.wpp.imageStorageSharePointDocumentLibrary ||
            props.wpp.imageStorageSharePointDocumentLibrary === ""
          }
          // hideStockImages
          hideLocalMultipleUploadTab
          hideOrganisationalAssetTab
          buttonLabel="Change image"
          onSave={onImagePickerSave}
          context={props.context}
          isPanelOpen={isImagePickerVisible}
          onCancel={() => setIsImagePickerVisible(false)}
        />

        {/*  */}
        <Image src={imageUrl} width={"100%"} height={"56.25%"} />

        {!props.wpp.imageStorageSharePointDocumentLibrary && (
          <MessageBar messageBarType={MessageBarType.error}>
            You must first specify a SharePoint Document Library in the web part
            property pane to use the Image FilePicker
          </MessageBar>
        )}
      </div>

      <div>
        <Text>URL Picker Demo</Text>

        <div
          style={{
            display: "flex",
            flexDirection: "row",
            gap: "20px",
            flexWrap: "wrap",
          }}
        >
          {urls.map(u => {
            return <Link href={u}>Link</Link>;
          })}
        </div>

        <FilePicker
          // accepts={["",".gif", ".jpg", ".jpeg", ".bmp", ".dib" , ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"]}
          hideLinkUploadTab
          // hideLocalUploadTab
          hideOneDriveTab
          hideWebSearchTab
          // hideStockImages
          hideLocalMultipleUploadTab
          hideOrganisationalAssetTab
          buttonLabel="Change URL's"
          onSave={onUrlPickerSave}
          context={props.context}
          isPanelOpen={isUrlPickerVisible}
          onCancel={() => setIsUrlPickerVisible(false)}
        />
      </div>
    </div>
  );
};
