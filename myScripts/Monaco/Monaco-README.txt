To update Monaco in the future:
1. In this folder, run "sd delete vs\...".
2. In a brand-new folder, run "npm install monaco-editor".  This will install it from https://www.npmjs.com/package/monaco-editor
3. Browse to "node_modules\monaco-editor\dev\vs" folder and copy the contents of that folder to a version-numbered folder here.  The version-numbering provides:
	a. A reference point for when Monaco was updated
	b. The ability to easily do "sd delete ..." on the old and "sd add ..." on the new -- and not have the deletions and additions conflict.
4. As per "b." above, sd delete the old files, and bulk add the new ones.
5. Do a bulk replacement of "Scripts/Monaco/<old-version-#>/vs" with "Scripts/Monaco/<NEW-version-#>/vs" in all HTML files under the RichApiAgaveWeb folder.
