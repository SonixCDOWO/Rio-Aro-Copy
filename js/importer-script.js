$(document).ready(function() {
    const fileInput = $('#excel-file-input');
    const previewContainer = $('#preview-container');
    const importBtn = $('#import-btn');
    const editModal = new bootstrap.Modal(document.getElementById('editPersonModal'));
    const modalFormFields = $('#modal-form-fields');
    let activePersonElement = null; // Guardará el elemento de la persona que se está editando

    fileInput.on('change', function(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            console.log("--- LOG: Datos leídos del Excel ---", jsonData);
            renderPreview(jsonData);
            importBtn.show();
        };
        reader.readAsArrayBuffer(file);
    });

    function renderPreview(people) {
        previewContainer.empty();
        if (people.length === 0) {
            previewContainer.html('<p class="text-muted">El archivo está vacío o no tiene el formato correcto.</p>');
            return;
        }

        // Agrupar por Comunidad -> Torre -> Casa
        const groupedData = people.reduce((acc, person, index) => {
            const com = person['COMUNIDAD'] || 'Sin Comunidad';
            const torre = person['TORRE'] || 'Sin Torre';
            const casa = person['CASA O APTO'] || 'Sin Casa';

            if (!acc[com]) acc[com] = {};
            if (!acc[com][torre]) acc[com][torre] = {};
            if (!acc[com][torre][casa]) acc[com][torre][casa] = [];

            person.__tempId = `person_${index}`; // ID temporal único
            acc[com][torre][casa].push(person);
            return acc;
        }, {});
        
        console.log("--- LOG: Datos agrupados ---", groupedData);
        
        let accordionHtml = '<div class="accordion">';
        Object.keys(groupedData).forEach((com, i) => {
            accordionHtml += `
                <div class="accordion-item">
                    <h2 class="accordion-header">
                        <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseCom${i}">
                            Comunidad: ${com}
                        </button>
                    </h2>
                    <div id="collapseCom${i}" class="accordion-collapse collapse">
                        <div class="accordion-body">
            `;
            Object.keys(groupedData[com]).forEach(torre => {
                accordionHtml += `<h5>Torre: ${torre}</h5>`;
                Object.keys(groupedData[com][torre]).forEach(casa => {
                    accordionHtml += `<p class="ms-3"><strong>Casa/Apto: ${casa}</strong></p><ul class="list-group ms-4 mb-3">`;
                    groupedData[com][torre][casa].forEach(person => {
                        // Guardamos todos los datos de la persona en atributos data-*
                        const dataAttributes = Object.entries(person).map(([key, value]) => `data-${key.toLowerCase().replace(/ /g, '-')}="${value}"`).join(' ');
                        accordionHtml += `<li class="list-group-item person-item" ${dataAttributes}>${person['APELLIDOS Y NOMBRES'] || 'Nombre no encontrado'}</li>`;
                    });
                    accordionHtml += `</ul>`;
                });
            });
            accordionHtml += `</div></div></div>`;
        });
        accordionHtml += `</div>`;
        previewContainer.html(accordionHtml);
    }

    // Evento para abrir el modal de edición
    previewContainer.on('click', '.person-item', function() {
        activePersonElement = $(this);
        modalFormFields.empty();
        const personData = activePersonElement.data();
        
        let formHtml = '';
        for (const key in personData) {
            // Convertir 'nombres-apellidos' de nuevo a 'Nombres Apellidos' para la etiqueta
            const label = key.replace(/-/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
            formHtml += `
                <div class="col-md-4">
                    <label class="form-label">${label}</label>
                    <input type="text" class="form-control" data-key="${key}" value="${personData[key]}">
                </div>
            `;
        }
        modalFormFields.html(formHtml);
        editModal.show();
    });

    // Evento para guardar los cambios del modal
    $('#save-modal-changes-btn').on('click', function() {
        if (!activePersonElement) return;

        modalFormFields.find('input').each(function() {
            const input = $(this);
            const key = input.data('key');
            const value = input.val();
            // Actualizar el atributo data-* en el elemento de la lista
            activePersonElement.attr(`data-${key}`, value);
        });
        
        // Actualizar el texto visible en la lista
        const displayName = activePersonElement.attr('data-apellidos-y-nombres');
        activePersonElement.text(displayName || 'Nombre no encontrado');

        editModal.hide();
        Toastify({ text: "Cambios guardados en la vista previa.", backgroundColor: "blue" }).showToast();
        activePersonElement = null;
    });
    
    // Evento para la importación final
    importBtn.on('click', function() {
        let finalPayload = [];
        $('.person-item').each(function() {
            const personElement = $(this);
            const personData = personElement.data();
            let cleanPerson = {};
            // Reconstruir el objeto con las claves originales del Excel
            for (const key in personData) {
                const originalKey = key.replace(/-/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
                cleanPerson[originalKey] = personData[key];
            }
            delete cleanPerson['Tempid']; // Eliminar el ID temporal
            finalPayload.push(cleanPerson);
        });

        console.log("--- LOG: Payload final a enviar al backend ---", finalPayload);

        fetch('/api/bulk-import', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ datos: finalPayload })
        })
        .then(res => {
            if (res.ok) {
                Toastify({ text: "¡Datos importados exitosamente!", duration: 5000, backgroundColor: "green" }).showToast();
                previewContainer.empty();
                importBtn.hide();
                fileInput.val(''); // Limpiar el input de archivo
            } else {
                Toastify({ text: "Error en la importación. Revisa la consola del servidor.", backgroundColor: "red" }).showToast();
            }
        });
    });
});