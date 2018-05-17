(function () {
    'use strict';

    // La función Office.initialize se debe definir para todas las páginas de la aplicación.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayDataOrRedirect();
        });
    };

    // Comprueba si existe un enlace y muestra la visualización,
    //        o redirige a la página Enlace de datos.
    function displayDataOrRedirect() {
        Office.context.document.bindings.getByIdAsync(
            app.bindingID,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var binding = result.value;
                    displayDataForBinding(binding);
                    binding.addHandlerAsync(
                        Office.EventType.BindingDataChanged,
                        function () { displayDataForBinding(binding); }
                    );
                } else {
                    window.location.href = '../DataBinding/DataBinding.html';
                }
            });
    }

    // Consulta sus datos al enlace y después delega en el script de visualización.
    function displayDataForBinding(binding) {
        binding.getDataAsync(
            {
                coercionType: Office.CoercionType.Table,
                valueFormat: Office.ValueFormat.Unformatted,
                filterType: Office.FilterType.OnlyVisible
            },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    visualization.display($('#data-display'), result.value, showError);
                } else {
                    showError('No se pueden leer datos.');
                }
            }
        );

        function showError(message) {
            $('#data-display').html(
                '<div class="notice">' +
                '    <h3>Error</h3>' + $('<p/>', { text: message })[0].outerHTML +
                '    <a href="../DataBinding/DataBinding.html">' +
                '        <b>¿Desea enlazar a otro intervalo de datos?</b>' +
                '    </a>' +
                '</div>');
        }
    }

})();
