# PnP SPFX FilePicker Control Sample

### 1 Introuction

- This very rough and ready sample illustrates how to use the FilePicker control in a SPFX webpart.
- It covers both selecting one or more URLs (for example, a SharePoint Site Page) and selecting Image URL's (including the ability to select an image from your own HDD)

### 2 Notes

- At time of writing, it is recommended to use V2.9.0 of the PNP SPFX controls until 3.11 lands - the FilePicker currently crashes the browser in V3 when trying to select a link

### 3 Running the sample

- Install Node v16.17.1
- npm i
- npm run serve (or gulp serve)

Make sure to specify the Document Library you want to save locally uploaded images to!
