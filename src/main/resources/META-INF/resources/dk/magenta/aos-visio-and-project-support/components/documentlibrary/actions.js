"use strict";

if (Alfresco !== undefined && Alfresco.doclib !== undefined && Alfresco.doclib.Actions !== undefined) {

    /**
     * Redefine the existing edit online actions
     */
    Alfresco.doclib.Actions.prototype.getProtocolForFileExtension = function(fileExtension) {
        var msProtocolNames =
        {
            'doc'  : 'ms-word',
            'docx' : 'ms-word',
            'docm' : 'ms-word',
            'dot'  : 'ms-word',
            'dotx' : 'ms-word',
            'dotm' : 'ms-word',
            'xls'  : 'ms-excel',
            'xlsx' : 'ms-excel',
            'xlsb' : 'ms-excel',
            'xlsm' : 'ms-excel',
            'xlt'  : 'ms-excel',
            'xltx' : 'ms-excel',
            'xltm' : 'ms-excel',
            'ppt'  : 'ms-powerpoint',
            'pptx' : 'ms-powerpoint',
            'pot'  : 'ms-powerpoint',
            'potx' : 'ms-powerpoint',
            'potm' : 'ms-powerpoint',
            'pptm' : 'ms-powerpoint',
            'pps'  : 'ms-powerpoint',
            'ppsx' : 'ms-powerpoint',
            'ppam' : 'ms-powerpoint',
            'ppsm' : 'ms-powerpoint',
            'sldx' : 'ms-powerpoint',
            'sldm' : 'ms-powerpoint',

            /**
             * Added support for project and visio
             */

            'mpp'  : 'ms-project',
            'vsd'  : 'ms-visio',
            'vsdx'  : 'ms-visio.drawing'
        };
        return msProtocolNames[fileExtension];
    };

    Alfresco.doclib.Actions.prototype.onlineEditMimetypes = {

        "application/msword": "Word.Document",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "Word.Document",
        "application/vnd.ms-word.document.macroenabled.12": "Word.Document",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template": "Word.Document",
        "application/vnd.ms-word.template.macroenabled.12": "Word.Document",

        "application/vnd.ms-powerpoint": "PowerPoint.Slide",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation": "PowerPoint.Slide",
        "application/vnd.ms-powerpoint.presentation.macroenabled.12": "PowerPoint.Slide",
        "application/vnd.openxmlformats-officedocument.presentationml.slideshow": "PowerPoint.Slide",
        "application/vnd.ms-powerpoint.slideshow.macroenabled.12": "PowerPoint.Slide",
        "application/vnd.openxmlformats-officedocument.presentationml.template": "PowerPoint.Slide",
        "application/vnd.ms-powerpoint.template.macroenabled.12": "PowerPoint.Slide",
        "application/vnd.ms-powerpoint.addin.macroenabled.12": "PowerPoint.Slide",
        "application/vnd.openxmlformats-officedocument.presentationml.slide": "PowerPoint.Slide",
        "application/vnd.ms-powerpoint.slide.macroEnabled.12": "PowerPoint.Slide",

        "application/vnd.ms-excel": "Excel.Sheet",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "Excel.Sheet",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template": "Excel.Sheet",
        "application/vnd.ms-excel.sheet.macroenabled.12": "Excel.Sheet",
        "application/vnd.ms-excel.template.macroenabled.12": "Excel.Sheet",
        "application/vnd.ms-excel.addin.macroenabled.12": "Excel.Sheet",
        "application/vnd.ms-excel.sheet.binary.macroenabled.12": "Excel.Sheet",
        "application/vnd.visio": "Visio.Drawing",
        "application/vnd.visio2013": "Visio.Drawing",
        "application/vnd.ms-project": "Microsoft.Project"
    };
}