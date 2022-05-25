var lista;
var Grilla;
var ayuda = [];
var tieneSubTotalExportar = "";
var _tab;

window.onload = function () {
    Http.get("Postulante/cargar", mostrarRpt);    
}

function mostrarRpt(rpta) {
    if (rpta) {
        lista = rpta.split(sepLista);
        var preIndice;
        var listaPostulante = lista[0].split(sepRegistros);
        preIndice = lista[1].split('~');
        _tab = preIndice[1];
        var indices = Array.from(preIndice[0].split(','), Number);
        var listaTipoDocum = lista[2].split(sepRegistros);
        var listaUniMinera = lista[3].split(sepRegistros);
        var listaEmpresas = lista[4].split(sepRegistros);
        var listaAreas = lista[5].split(sepRegistros);
        var listaPuestos = lista[6].split(sepRegistros);
        var listaEstados = lista[7].split(sepRegistros);
        var tipoIndice = ['hdn', 'cbo', 'nro', 'txt', 'fec', 'cbo', 'cbo', 'cbo', 'cbs', 'cbo'];
        var listaMaxlenth = ['', '', '9', '200', '10', '', '', '', '', ''];

        ayuda.push(listaTipoDocum, listaUniMinera, listaEmpresas, listaAreas, listaPuestos, listaEstados);

        Grilla = new GUI.Grilla(grillaPostulante, listaPostulante, 'grillaPostulante', null, "Total de Registros: ", indices,
            ayuda, null, null, 9, 10, true, true, true, true, true, true, null, 'EDITABLE', null, null, null, tipoIndice, null, false, listaMaxlenth);
    }
}

function nuevoRegistro() {
    var matriz = Grilla.ObtenerMatriz();
    var regMatriz = matriz.length;
    var fila = ['', '', '', '', '', '', '', '', '', ''];
    matriz.push(fila);
    Grilla.AgregarNuevo();
    altert('NUEVO REGISTRO');
}
function editarRegistro(id, cod, fila) {
    var btnEdit = fila.childNodes[10].firstChild;
    var btnDelete = fila.childNodes[11].firstChild;
    var flagFila = fila.getAttribute("flag-accion");
    var listaControles = fila.getElementsByClassName("GE");
    Validacion.ValidarNumerosEnLinea("N", fila);

    if (flagFila == null || flagFila == "") {
        fila.setAttribute("flag-accion", "EDITAR");
        btnEdit.src = hdfRaiz.value + "Images/save.svg";
        btnEdit.title = 'Guardar';
        btnDelete.src = hdfRaiz.value + "Images/return.svg";
        btnDelete.title = 'volver';
        habilitarControles(true, listaControles);
    }
    if (flagFila == "EDITAR") {
        //NOTA: para quitar la clase requerido
        var txtNombres = fila.childNodes[3].firstChild;
        txtNombres.classList.remove("GE");


        if (Validacion.ValidarRequeridos("GE", fila) == 0) return;
        if (Validacion.ValidarNumeros("N", fila) == 0) return;

        fila.setAttribute("flag-accion", "");
        btnEdit.src = hdfRaiz.value + "Images/edit.svg";
        btnEdit.title = 'Editar';
        btnDelete.src = hdfRaiz.value + "Images/delete.svg";
        btnDelete.title = 'Eliminar';

        alert('SE GRABARON LOS DATOS EN CUESTION :');
        //NOTA: para devolverle la clase requerido
        txtNombres.classList.add("GE");


        habilitarControles(false, listaControles);
    }
}
function eliminarRegistro(id, cod, fila) {
    var btnEdit = fila.childNodes[10].firstChild;
    var btnDelete = fila.childNodes[11].firstChild;
    var flagFila = fila.getAttribute("flag-accion");
    var listaControles = fila.getElementsByClassName("GE");

    if (flagFila == "EDITAR") {
        var nroCtl = listaControles.length;
        var ctlData;
        for (var i = 0; i < nroCtl; i++) {
            ctlData = listaControles[i];
            if (ctlData.type == "hidden")ctlData.SetValor(ctlData.getAttribute("data-val"));
            else ctlData.value = ctlData.getAttribute("data-val");
        }

        fila.setAttribute("flag-accion", "");
        btnEdit.src = hdfRaiz.value + "Images/edit.svg";
        btnEdit.title = 'Editar';
        btnDelete.src = hdfRaiz.value + "Images/delete.svg";
        btnDelete.title = 'Eliminar';
        habilitarControles(false, listaControles);
    }
    if (flagFila == null || flagFila == "") {
        alert('SE ELIMINARA LA FILA SELECCIONADA: ');
    }
}
function habilitarControles(habilitado, lista) {
    var ncontroles = lista.length;
    for (var i = 0; i < ncontroles; i++) {
        lista[i].disabled = !habilitado;
        !habilitado == false ? lista[i].classList.remove("colorDisabled") : lista[i].classList.add("colorDisabled");
    }
}


function iniciarComboRegistro(idGrilla) {
}
function seleccionarFila() {
}
function iniciarComboSearchRegistro(idGrilla) {
}
function pulsarBotonGrilla(idGrilla, input, indiceCol, indiceRow, matriz) {
}
function cambiarComboRegistro(idGrilla, combo, indiceCol, indiceRow, matriz) {
}
function cambiarInputRegistro(idGrilla, input, indiceCol, indiceRow, matriz) {
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

//if (ctlData.className.indexOf("InputcboSeacrh") > -1) && ctlData.tagName=="INPUT"

*/

/*
NOTA: PARA REMOVER UN REQUERIDO
=================================
var txtTiempo = fila.childNodes[1].firstChild;
txtTiempo.classList.remove("GE");
*/

