"use strict";
module.exports = {
    src: {
        js: [
            "src/sharepoint/**/*.app.js",
            "src/sharepoint/**/*.module.js",
            "src/sharepoint/**/*.js"
        ],
        css: [
            "src/css/sharepoint/**/*.css"
        ],
        html: [
            "src/sharepoint/**/*.html"
        ],
        images: [
            "src/sharepoint/images/**/*"
        ]
    },

    // Common is only used in STANDALONE!
    common: {
        js: [
            "bower_components/jquery/dist/jquery.min.js",
            "bower_components/jquery.ui/ui/core.js",
            "bower_components/jquery.ui/ui/widget.js",
            "bower_components/jquery.ui/ui/mouse.js",
            "bower_components/jquery.ui/ui/draggable.js",
            "bower_components/jquery.ui/ui/droppable.js",
            "bower_components/jquery.ui/ui/resizable.js"
        ],
        css: []
    },
    resources: {
        i18n: [
            "src/resources/i18n/sharepoint/**/*.xml"
        ]
    }
};
