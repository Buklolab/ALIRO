﻿/* Funcionalidad de aplicación común */

var app = (function () {
    'use strict';

    var app = {};

    app.bindingID = 'myBinding';

    // Función de inicialización común (se llama desde todas las páginas).
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });


        // Después de la inicialización, exponga una función de notificación común.
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };
    };

    return app;
})();