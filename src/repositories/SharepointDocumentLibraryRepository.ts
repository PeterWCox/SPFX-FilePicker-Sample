import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { IFilePickerResult } from "@pnp/spfx-controls-react";
import { SharePointRepository } from "./SharePointRepository";
import { IFileAddResult } from "@pnp/sp/files";

export class SharePointDocumentLibraryRepository extends SharePointRepository {
  // e.g. file = output of the PnP File Picker control onChange handler
  // e.g. relativeUrl = "/sites/LightspeedDemo/Shared Documents"
  public async uploadFileToSharePointDocumentLibrary(
    file: IFilePickerResult,
    documentLibrary: string
  ): Promise<IFileAddResult> {
    try {
      //Download the image
      const downloadedFile = await file.downloadFileContent();

      //Make sure we the filename is unique
      const uniqueFileName = `${file.fileName}_${new Date().getTime()}`;

      let result: IFileAddResult;

      if (file.fileSize <= 10485760) {
        result = await this._sp.web
          .getFolderByServerRelativePath(documentLibrary)
          .files.addUsingPath(uniqueFileName, downloadedFile, {
            Overwrite: true,
          });
      } else {
        result = await this._sp.web
          .getFolderByServerRelativePath(documentLibrary)
          .files.addChunked(
            uniqueFileName,
            file,
            (data: any) => {
              console.log(`progress`);
            },
            true
          );
      }

      return result;
    } catch (error) {
      console.error(
        "An error has occured uploading a locally uploaded image to a document library",
        error
      );
    }
  }
}
