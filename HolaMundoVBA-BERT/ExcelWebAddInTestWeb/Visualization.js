var visualization = (function () {
    'use strict';

    var visualization = {};

    // Genera y devuelve un objeto Office.TableData con datos de ejemplo.
    visualization.generateSampleData = function () {
        var sampleHeaders = [['Nombre', 'Grado']];
        var sampleRows = [
            ['Benjamín', 79],
            ['Alejandra', 95],
            ['Jacobo', 86],
            ['Ernesto', 93]];
        return new Office.TableData(sampleRows, sampleHeaders);
    };

    // Muestra una visualización basada en los parámetros siguientes:
    //        $element: elemento jQuery donde se mostrará la visualización.
    //        data: objeto Office.TableData que contiene los datos.
    //        errorHandler: devolución de llamada de error que acepta una descripción de cadena.
    visualization.display = function ($element, data, errorHandler) {
        if (data.rows.length < 1 || data.rows[0].length < 2) {
            errorHandler('El intervalo de datos debe contener al menos 1 fila y al menos 2 columnas.');
            return;
        }

        var maxBarWidthInPixels = 200;
        var $table = $('<table class="visualization" />');

        if (data.headers !== null && data.headers.length > 0) {
            var $headerRow = $('<tr />').appendTo($table);
            $('<th />').text(data.headers[0][0]).appendTo($headerRow);
            $('<th />').text(data.headers[0][1]).appendTo($headerRow);
        }

        for (var i = 0; i < data.rows.length; i++) {
            var $row = $('<tr />').appendTo($table);
            var $column1 = $('<td />').appendTo($row);
            var $column2 = $('<td />').appendTo($row);

            $column1.text(data.rows[i][0]);
            var value = data.rows[i][1];
            var width = maxBarWidthInPixels * value / 100.0;
            var $visualizationBar = $('<div />').appendTo($column2);
            $visualizationBar.addClass('bar').width(width).text(value);
        }

        $element.html($table[0].outerHTML);
    };

    return visualization;
})();
