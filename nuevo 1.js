
var listaControles = fila.getElementsByClassName("GE");
habilitarControles(true, listaControles);



::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
function habilitarControles(habilitado, lista) {
    var ncontroles = lista.length;
    for (var i = 0; i < ncontroles; i++) {
        lista[i].disabled = !habilitado;
        !habilitado == false ? lista[i].classList.remove("colorDisabled") : lista[i].classList.add("colorDisabled");
    }
}




function convertirMayuscula() {
    var listaControles = document.getElementsByClassName('E');
    var nlistaControles = listaControles.length;
    for (var i = 0; i < nlistaControles; i++) {
        listaControles[i].classList.add('Upper');
    }
}

var dato ="1,2,3,4,5";
var salida = Array.from(dato.split(','),Number);


