var lista;
var Grilla;
var ayuda = [];
var tieneSubTotalExportar = "";

window.onload = function () {
    Http.get("Postulante/cargar", mostrarRpt);
}

function mostrarRpt(rpta) {
    if (rpta) {
        lista = rpta.split(sepLista);
        var indices = [1, 5, 6, 7, 8, 9];
        var listaPostulante = lista[0].split(sepRegistros);
        var listaTipoDocum = lista[2].split(sepRegistros);
        var listaUniMinera = lista[3].split(sepRegistros);
        var listaEmpresas = lista[4].split(sepRegistros);
        var listaAreas = lista[5].split(sepRegistros);
        var listaPuestos = lista[6].split(sepRegistros);
        var listaEstados = lista[7].split(sepRegistros);
        var tipoIndice = ['hdn', 'cbo', 'txt', 'txt', 'txt', 'cbo', 'cbo', 'cbo', 'cbs', 'cbo'];

        ayuda.push(listaTipoDocum, listaUniMinera, listaEmpresas, listaAreas, listaPuestos, listaEstados);

        Grilla = new GUI.Grilla(grillaPostulante, listaPostulante, 'grillaPostulante', null, "Total de Registros: ", indices,
        ayuda, null, null, 9, 10, true, true, true, true, true, true, null, 'EDITABLE', null, null, null, tipoIndice, null, false, false);
    }
}


function iniciarComboRegistro(idGrilla) {
}
function seleccionarFila() {
}
function iniciarComboSearchRegistro(idGrilla) {
}
function cambiarComboRegistro(idGrilla, combo, indiceCol, indiceRow, matriz) {
}
function nuevoRegistro() {
    alert('nuevo');
}
function editarRegistro(id, cod, fila) {
    var btnEdit = fila.childNodes[10].firstChild;
    var btnDelete = fila.childNodes[11].firstChild;
    var flagFila = fila.getAttribute("flag-accion");

    if (flagFila == null || flagFila == "") {
        fila.setAttribute("flag-accion", "EDITAR");
        btnEdit.src = hdfRaiz.value + "Images/save.svg";
        btnEdit.title = 'Guardar';
        btnDelete.src = hdfRaiz.value + "Images/return.svg";
        btnDelete.title = 'volver';
    }
    if (flagFila == "EDITAR") {
        fila.setAttribute("flag-accion", "");
        btnEdit.src = hdfRaiz.value + "Images/edit.svg";
        btnEdit.title = 'Editar';
        btnDelete.src = hdfRaiz.value + "Images/delete.svg";
        btnDelete.title = 'Eliminar';
        alert('SE GRABARON LOS DATOS EN CUESTION :');
    }    
}
function eliminarRegistro(id, cod, fila) {
    var btnEdit = fila.childNodes[10].firstChild;
    var btnDelete = fila.childNodes[11].firstChild;
    var flagFila = fila.getAttribute("flag-accion");

    if (flagFila == "EDITAR") {
        fila.setAttribute("flag-accion", "");
        btnEdit.src = hdfRaiz.value + "Images/edit.svg";
        btnEdit.title = 'Editar';
        btnDelete.src = hdfRaiz.value + "Images/delete.svg";
        btnDelete.title = 'Eliminar';
    }
    if (flagFila == null || flagFila == "") {
        alert('SE ELIMINARA LA FILA SELECCIONADA: ');
    }
}

/*
fila.setAttribute("flag-accion", "ELIMINAR"); // AGREGAR EDITAR NUEVO
btnEdit.src = hdfRaiz.value + "Images/save.svg";
btnEdit.title = 'Guardar';
btnDelete.src = hdfRaiz.value + "Images/return.svg";
btnDelete.title = 'volver';


btnEdit.src = hdfRaiz.value + "Images/edit.svg";
btnEdit.title = 'Editar';
btnDelete.src = hdfRaiz.value + "Images/delete.svg";
btnDelete.title = 'Eliminar';
habilitarBotonesGrilla();
deshabilitarBotonesGrilla(indiceFila);

fila.childNodes[10].firstChild.style.pointerEvents = "none";
btnNuevogrillaRegTiempo.style.pointerEvents = "auto";

habilitarControles(true, lista);
*/

