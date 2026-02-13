Diagram Frame SharePoint Visio Web Part
---------------------------------------

* * *

This package contains a SharePoint Framework (SPFx) Web Part that can also be configured as a full-page application and integrated as a Microsoft Teams tab.

The web part allows embedding Microsoft Visio diagrams into modern SharePoint pages, full-page applications, and Teams tabs. It supports viewing and interacting with Visio diagrams while providing enhanced navigation and customization options compared to the default SharePoint “File Viewer” web part.

One of the key capabilities is the ability to navigate through diagram hierarchies, enabling users to drill down into linked diagrams and individual pages within those diagrams. This makes it easier to explore complex processes or organizational charts that are built across multiple interconnected Visio files.

Additional functionality includes setting which Visio page is displayed initially and controlling the default zoom level to optimize how the diagram is presented when the web part loads. You can also specify exact dimensions—either relative (percentage) or fixed sizes—for the diagram container, which helps integrate the diagrams smoothly with varying page layouts and screen sizes.

The web part provides an interface tailored specifically for Visio files with a dedicated file browser, improving the ease of selecting and managing diagrams directly within SharePoint.

Several display options are available to customize the user experience. For example, you can hide interface components such as the diagram toolbar, boundaries, or border to create a cleaner, more focused view. Similarly, interactive features like hyperlinks, zooming, panning, and the navigation pane can be disabled if desired, giving you control over how users interact with the diagrams.

The solution also supports rendering as a full-width web part or as a full page (application page mode). This flexibility allows scenarios that require immersive, large-scale diagram presentations, making it suitable for dashboards or detailed process views.

### File Browser

The web part includes a file browser designed specifically for Visio diagrams, allowing users to easily browse and select files stored within SharePoint.

### Start Page

You can configure the web part to open a specific page within a Visio diagram by name. If the designated page is not available, it will default to displaying the first page in the diagram.

### Select Default Zoom Level

The default zoom level of the diagram can be set to control the initial scale at which the diagram appears.

### Hierarchy Navigation

Supports hierarchical navigation, allowing users to drill down through multiple levels of linked diagrams and pages for an organized, step-by-step exploration of complex Visio content.

### Full Width and Application Page Mode Support

The web part can be configured to occupy the full width of a page or be used in application page mode to provide a dedicated and uninterrupted Visio diagram viewing experience.

### Support for Custom Width/Height Settings

The dimensions of the diagram container can be customized using percentage values or fixed measurements, balancing responsiveness with layout precision.

### Support for Disabling Menu Items

Options are available to hide the diagram toolbar, boundary, and border, enabling a streamlined visual presentation.

### Support for Disabling Some Interactivity

Interactivity features such as hyperlinks, zoom, pan, and the navigator can be selectively disabled to suit different use cases or security requirements.

What's New in Version 1.1
-------------------------

*   Enabled use of the web part as a full-width page
*   Enabled embedding the web part as a Teams tab
*   Optimized flickering on first load
*   Improved experience with file upload (uploads default to the site document library)

What's New in Version 1.2
-------------------------

*   Added German localization
*   Introduced experimental support for navigation
*   Switched to default CDN hosting with jsdelivr

What's New in Version 1.3
-------------------------

*   Added option to open hyperlinks in a separate browser window
*   Added option to open Office file links inline (using `?web=1` query parameter) instead of triggering downloads
*   Improved handling to reduce flickering when the diagram first loads
*   Added an optional navigation header to provide additional context during diagram navigation

More Information
----------------

For further details, including setup instructions and source code, see the project site:

[https://unmanagedvisio.com/products/diagram-frame-sharepoint-visio-web-part/](https://unmanagedvisio.com/products/diagram-frame-sharepoint-visio-web-part/)
