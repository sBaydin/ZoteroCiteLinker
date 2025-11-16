# ZoteroCiteLinker

An MS Word Add-in that creates an interactive bridge between your in-text Zotero citations and your bibliography, all from a convenient custom Ribbon tab.

This add-in scans your document for Zotero-generated citation fields (e.g., `[1]`, `(Author, 2020)`) and intelligently links them to the corresponding entries in your bibliography, allowing for instant navigation.

[ss.png]

## 🌟 Features

* **One-Click Linking:** Automatically link all citations in the document.
* **Step-by-Step Linking:** A manual "debug" mode to process citations one by one.
* **Smart Unlinking:** Safely remove *only* the links created by this add-in, without touching your Table of Contents or other hyperlinks.
* **Utility Functions:** Activate (or separately remove) all web URLs and email addresses in your document.
* **Customization:** Change the color of all Zotero citations or apply a custom Word "Character Style" for advanced formatting.
* **Pre-flight Check:** Verifies that the Zotero application is running before attempting to link.
* **Compatibility:** Works with over 15 common Zotero citation styles (listed in the "How to Use" section below).

## 💾 Installation Instructions

This add-in is installed by placing the `.dotm` file into Word's trusted STARTUP folder. This will automatically load the "Zotero Linker" Ribbon tab every time you open Word.

1.  **Get the File:**
    * Download the latest `ZoteroCiteLinker.dotm` file from the [Releases page](httpsTo-your-releases-page) on this GitHub repository.

2.  **Find your Word STARTUP Folder:**
    * Open MS Word.
    * Go to **File** -> **Options** -> **Trust Center**.
    * Click the **Trust Center Settings...** button.
    * Go to the **Trusted Locations** tab.
    * In the list, find the location with the description "**User STARTUP**" (veya Türkçe Word'de "**Kullanıcı Başlangıç**").
    * **Click on it and copy the full path** from the bottom of the window.

    [Buraya Güven Merkezi'ndeki (Trust Center) Başlangıç (STARTUP) klasör yolunu gösteren bir ekran görüntüsü ekleyin]

    * The path will look something like this:
        `C:\Users\<YourUserName>\AppData\Roaming\Microsoft\Word\STARTUP`

3.  **Copy the Add-in File:**
    * Open Windows File Explorer.
    * Paste the path you just copied into the address bar and press **Enter**.
    * Copy the `ZoteroCiteLinker.dotm` file into this `STARTUP` folder.

4.  **Restart Word & Enable Content:**
    * Close MS Word completely and re-open it.
    * When Word starts, it will likely show a yellow **"SECURITY WARNING"** bar at the top, stating "Macros have been disabled."
    * You **must** click the **"Enable Content"** button. This tells Word to trust the add-in.
    * You should now see a new **"Zotero Linker"** tab on your Word Ribbon.

## 📖 How to Use: The Ribbon Guide

[Buraya Zotero Linker Şeridindeki (Ribbon) tüm düğmeleri gösteren bir resim ekleyin]

Here is a breakdown of each button on the "Zotero Linker" tab.

### Main Functions

* **Link All Citations (`ZCL_LinkCitationAll`)**
    * This is the main button. It first unlinks all previous citations and then automatically scans and links every Zotero citation in your document. A message will appear showing how many links were created.
    * **Requirement:** The Zotero application must be running for this to work.

* **Link Selectively (`ZCL_LinkCitationSelect`)**
    * This is a debug/manual mode. It will go through the document one citation field at a time and ask you (Yes/No/Cancel) if you want to process that specific group.
    * **Requirement:** The Zotero application must be running.

* **Unlink Citations (`ZCL_UnlinkCitations`)**
    * Safely removes *only* the hyperlinks from your Zotero citations (bookmarks starting with `CITE_`). It also automatically resets their color back to black (Auto).
    * This **does not** touch other hyperlinks in your document (like web links or your Table of Contents).

### Utility Functions

* **Activate URLs (in Bib) (`ZCL_LinkBibliographyURLs`)**
    * This utility scans *only* your Zotero Bibliography section and converts all text URLs (e.g., `http://...`) into active, clickable hyperlinks.

* **Activate Emails (All) (`ZCL_LinkEmails`)**
    * This utility scans your *entire document* (not just the bibliography) and converts all email addresses (e.g., `name@domain.com`) into active `mailto:` links.

* **Unlink Web/Mail (`ZCL_UnlinkURLsMail`)**
    * The reverse of the above. This removes all `http://` and `mailto:` hyperlinks from the entire document.
    * This **does not** touch your Zotero citation links.

### Formatting & Help

* **Change Color (`ZCL_ChangeColor`)**
    * Opens a simple menu asking you to pick a color by number (Blue, Red, Auto/Black, etc.). This will apply your chosen color to all Zotero citation fields in the document.

* **Set Style (`ZCL_SetCitationStyle`)**
    * (Advanced) Allows you to apply a specific Word "Character Style" (e.g., "Emphasis") to your citations for custom formatting. A warning will appear, as applying a "Paragraph Style" by mistake can reformat your document.

* **Info (`ZCL_Information`)**
    * Displays a pop-up box listing all 17 citation styles supported by the add-in's engine:

| Proper Name | Zotero Style ID |
|---|---|
| American Chemical Society (ACS) | `american-chemical-society` |
| American Medical Association (AMA) | `american-medical-association` |
| American Political Science Ass. | `american-political-science-association` |
| American Psychological Ass. (APA) | `apa` |
| American Sociological Ass. (ASA) | `american-sociological-association` |
| Archives of Comp. Methods in Eng. | `archives-of-computational-methods-in-engineering` |
| BMC Medicine | `bmc-medicine` |
| Chicago Manual of Style (Author-Date) | `chicago-author-date` |
| China National Standard (Author-Date) | `china-national-standard-gb-t-7714-2015-author-date` |
| China National Standard (Numeric) | `china-national-standard-gb-t-7714-2015-numeric` |
| Elsevier - Harvard | `elsevier-harvard` |
| Harvard (Cite Them Right) | `harvard-cite-them-right` |
| IEEE | `ieee` |
| Modern Language Association (MLA) | `modern-language-association` |
| Molecular Plant | `molecular-plant` |
| Nature | `nature` |
| Vancouver | `vancouver` |

## 🤝 Acknowledgments

This add-in was built by modifying the excellent foundational code from:
* [altairwei/ZoteroLinkCitation](https://github.com/altairwei/ZoteroLinkCitation)
* [8gengen8/ZoteroCitationLink-lgg](https://github.com/8gengen8/ZoteroCitationLink-lgg)

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
