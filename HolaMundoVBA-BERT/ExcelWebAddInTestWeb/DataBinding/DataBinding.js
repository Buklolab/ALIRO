(function () {
    'use strict';

    // La función Office.initialize se debe definir para todas las páginas de la aplicación.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#bind-to-existing-data').click(bindToExistingData);

            if (dataInsertionSupported()) {
                $('#insert-sample-data').show();
                $('#insert-sample-data').click(insertSampleData);
            }
        });
    };

    // Enlaza la visualización a datos existentes.
    function bindToExistingData() {
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Table,
            { id: app.bindingID, sampleData: visualization.generateSampleData() },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    window.location.href = '../Home/Home.html';
                } else {
                    app.showNotification(result.error.name, result.error.message);
                }
            }
        );
    }

    // Comprueba si la aplicación actual admite establecer datos seleccionados.
    function dataInsertionSupported() {
        return Office.context.document.setSelectedDataAsync &&
            (Office.context.document.bindings &&
            Office.context.document.bindings.addFromSelectionAsync);
    }

    // Inserta datos de ejemplo en la selección actual (si se admite).
    function insertSampleData() {
        Office.context.document.setSelectedDataAsync(visualization.generateSampleData(),
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    Office.context.document.bindings.addFromSelectionAsync(
                        Office.BindingType.Table, { id: app.bindingID },
                        function (result) {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                window.location.href = '../Home/Home.html';
                            } else {
                                app.showNotification(result.error.name, result.error.message);
                            }
                        }
                    );
                } else {
                    app.showNotification(result.error.name, result.error.message);
                }
            }
        );
    }
})();
