history.pushState(null, null, document.URL);
window.addEventListener('popstate', function () {
    history.pushState(null, null, document.URL);
});
var sepCampos = '|';
var sepRegistros = '~';
var sepLista = '^';
var sepGrupoLista = '>|<';
var sepComodin = '=';

var contadorSessionActiva = 0;
var contadorSessionLogout = 0;
var sessionActiva = setInterval(function () {
    contadorSessionActiva = contadorSessionActiva + 1;
    contadorSessionLogout = contadorSessionLogout + 1;
    console.log("SessionActiva :" +contadorSessionActiva);
    console.log("SessionLogout :"+contadorSessionLogout);
    //if (contadorSessionLogout == 10800) {
    if (contadorSessionLogout == 10800) {
        clearInterval(sessionActiva);
        var url = hdfRaiz.value + 'Sistema/Logout';
        window.location.href  = url;
    }
    if (contadorSessionActiva==600) {
        Http.get("Sistema/obtenerVersion", mostrarVersion);
        function mostrarVersion(rpta) {
            if (validaResponseData(rpta)) {
                console.log(rpta);
                contadorSessionActiva = 0;
            }
        }
    }
}, 1000);

var Http = (function () {
    function Http() {
    }
    Http.get = function (url, callBack) {
        requestServer(url, "get", callBack);
    }
    Http.getBytes = function (url, callBack) {
        requestServer(url, "get", callBack, null, "arraybuffer");
    }
    Http.post = function (url, callBack, data) {
        requestServer(url, "post", callBack, data);
    }
    Http.postDownload = function (url, callBack, data) {
        requestServer(url, "post", callBack, data, "arraybuffer");
    }
    function requestServer(url, metodoHttp, callBack, data, tipoRpta) {
        var token = window.sessionStorage.getItem("token");
        var xhr = new XMLHttpRequest();
        xhr.open(metodoHttp, hdfRaiz.value + url);
        xhr.setRequestHeader("token", token);
        if (tipoRpta != null) xhr.responseType = tipoRpta;
        var contador = 0;

        var intervalo = setInterval(function () {
            contador = contador + 10;
            if (contador == 400) {
                mostrarloader();
            }
            if (xhr.status == 200 && xhr.readyState == 4) {
                clearInterval(intervalo);
                console.log(xhr.readyState);
                console.log(contador);
            }
        }, 10);

        xhr.onreadystatechange = function () {
            if (xhr.status == 200 && xhr.readyState == 4) {
                if (contador >= 400) {
                    cerrarloader();
                }
                if (tipoRpta != null) callBack(xhr.response);
                else callBack(xhr.responseText);
            }
        }
        if (data != null) xhr.send(data);
        else xhr.send();
       
    }
    return Http;
})();

var GUI = (function () {
    function GUI() {
    }
    GUI.Combo = function (cbo, lista, primerItem,tieneItemCompuesto) {
        var html = "";
        if (primerItem != null) {
            html += "<option value=''>";
            html += primerItem;
            html += "</option>";
        }
        var nRegistros = lista.length;
        var campos = [];
        for (var i = 0; i < nRegistros; i++) {
            campos = lista[i].split("|");
            html += "<option value='";
            html += campos[0];
            html += "'>";
            if (tieneItemCompuesto == true) {
                html += campos[0];
                html += "-";
                html += campos[1];
            } else {
                html += campos[1];
            }
            html += "</option>";
        }
        cbo.innerHTML = html;
    }

    //agregado por diego
    GUI.ComboSeacrh = function (cboSearch, lista, texto = "SELECCIONE") {

        var nlista = lista.length;   
        var campos = [];
        var liarray = [];
        var nliarray;
        var nombre = "";
        var matriz = [];
        var matrizFiltra = [];
        var nmatriz = 0;
        
        html = "";
        var divContent = cboSearch.parentNode.parentNode.parentNode;
        var button = divContent.children[0].children[0].children[1];
        var spanText = divContent.children[0].children[0].children[0];
        var inputhidden = divContent.children[0].children[0].children[2];
        var contentlista = divContent.children[0].children[1];
        var listaul = divContent.children[0].children[1].children[1];
        var input = divContent.children[0].children[1].children[0];
        llenarlista();
        if (texto == "") texto = "SELECCIONE";
        spanText.innerHTML = texto;
       
        divContent.onmouseleave = function () {
            if (contentlista.style.display == 'flex') {
                togleseacrh();
            }
            
        }
       
        function llenarlista() {
            matriz = [];
            for (var i = 0; i < nlista; i++) {
                campos = lista[i].split('|');
                matriz[i] = campos;
            }
            nmatriz = matriz.length;
            var contador = nmatriz > 300 ? 300 : nmatriz;

            for (var i = 0; i < contador; i++) {
                html += "<li data-item='";
                html += matriz[i][0];
                html += "' class='seacrh_option_list'>";
                html += matriz[i][1];
                html += "</li>";
            }
            listaul.innerHTML = html;
     
            configurarcontroles();
        }
        button.onclick = function () {
            togleseacrh();
        }

        spanText.onclick = function () {
            togleseacrh();
        }

        input.onkeyup = function () {
            var valor = input.value.toLowerCase();
            var html = "";

            for (var i = 0; i < nmatriz; i++) {
                //var registro = liarray[i].innerHTML;
                //if (registro.toLowerCase().indexOf(valor) > -1) {
                //    liarray[i].style.display = "block";
                //}
                //else {
                //    liarray[i].style.display = "none";
                //}
                if (matriz[i][1].toLowerCase().indexOf(valor) > -1) {                
                    matrizFiltra.push(matriz[i]);
                }
            }

            var contador = matrizFiltra.length > 300 ? 300 : matrizFiltra.length;
            for (var i = 0; i < contador; i++) {
                html += "<li data-item='";
                html += matrizFiltra[i][0];
                html += "' class='seacrh_option_list'>";
                html += matrizFiltra[i][1];
                html += "</li>";
            }
            matrizFiltra = [];
            listaul.innerHTML = html;
            configurarcontroles();
        }

        listaul.onscroll = function(UiEvent){
            console.log(UiEvent.detail);
        }

        function configurarcontroles() {
            liarray = divContent.getElementsByClassName('seacrh_option_list');
            nliarray = liarray.length;
            for (var i = 0; i < nliarray; i++) {
                liarray[i].onclick = function () {
                    spanText.innerHTML = this.innerHTML;
                    inputhidden.value = this.getAttribute('data-item');
                    itemselecionado = this.innerHTML;
                    nombre = this.innerHTML;
                    togleseacrh();
                    input.value = "";
                    inputhidden.onchange();
                }
            }
        }

        function togleseacrh() {
            if (inputhidden.disabled == true) return;
            if (button.innerHTML == "▲") {
                button.innerHTML = "▼";
                contentlista.style.display = 'flex';
            } else {
                button.innerHTML = "▲";
                contentlista.style.display = 'none';
            }
        }

        this.obtenerNombre = function(){
            return nombre;
        }

        this.setValor = function (valor) {
            var item;
            for (var i = 0; i < nmatriz; i++) {
                if (matriz[i][0] == valor) {
                    item = matriz[i];
                    break;
                }
            }
            if (item != null) {
                spanText.innerHTML = item[1];
                inputhidden.value = item[0];
                nombre = item[1];
                input.value = "";
            }       
        }

        this.bloquearCombo = function () {
            divContent.disabled = true;
             button.disabled = true;
             spanText.disabled = true;
             inputhidden.disabled = true;
             input.disabled = true;
        }

        if (texto != "SELECCIONE") {
            this.setValor(texto);
        }

        this.setTextoPorDefecto = function () {
            spanText.innerHTML = "SELECCIONE";
            inputhidden.value = "";
        }

        this.getTextoPorDefecto = function () {
            return spanText.innerHTML;
        }

    }

    GUI.Grilla = function (div, lista, id, botones, mensajeRegistros, indices, ayudas, tieneCheck, subtotales, registrosPagina, paginasBloque,
        tieneFiltros, tieneExportacion, tieneImprimir, tieneNuevoEditar, tieneEliminar, tieneCodigoOculto, listacheck,
        tipoGrilla, tieneCheckCabecera, listaColorFila, listaColorTextoFila, indiceTipos, tieneCalculofiltroFueraGrilla,
        listaBotonesAdicionales, listaMaxlenth) {
        
        var matriz = [];
        var nRegistros = lista.length;
        var nCampos;
        botones = (botones == null ? [] : botones);
        var nBotones = botones.length;
        var filaActual = null;
        var tipos = [];
        indices = (indices == null ? [] : indices);
        ayudas = (ayudas == null ? [] : ayudas);
        tieneCheck = (tieneCheck == null ? false : tieneCheck);
        subtotales = (subtotales == null ? [] : subtotales);
        registrosPagina = (registrosPagina == null ? 20 : registrosPagina);
        paginasBloque = (paginasBloque == null ? 10 : paginasBloque);
        tieneFiltros = (tieneFiltros == null ? false : tieneFiltros);
        tieneExportacion = (tieneExportacion == null ? false : tieneExportacion);
        tieneImprimir = (tieneImprimir == null ? false : tieneImprimir);
        tieneNuevoEditar = (tieneNuevoEditar = null ? false : tieneNuevoEditar);
        tieneEliminar = (tieneEliminar = null ? false : tieneEliminar);
        tieneCodigoOculto = (tieneCodigoOculto = null ? false : tieneCodigoOculto);
        tipoGrilla = (tipoGrilla == null ? 'NORMAL' : tipoGrilla);
        listaMaxlenth = (listaMaxlenth == null ? false : listaMaxlenth);
        tieneCheckCabecera = (tieneCheckCabecera == null ? false : tieneCheckCabecera);
        tieneCalculofiltroFueraGrilla = (tieneCalculofiltroFueraGrilla == null ? false : tieneCalculofiltroFueraGrilla);

        listaColorFila = (listaColorFila == null ? false : listaColorFila);
        //VALORES A TENER EN CUENTA
        //INDICE|NOMBRE_VALOR|COLOR
        listaColorTextoFila = (listaColorTextoFila == null ? false : listaColorTextoFila);
         //VALORES A TENER EN CUENTA
        //INDICE|NOMBRE_VALOR|COLOR

        indiceTipos = (indiceTipos == null ? [] : indiceTipos);
        listaBotonesAdicionales = (listaBotonesAdicionales == null ? [] : listaBotonesAdicionales);

        var nlistaBotonesAdicionales = listaBotonesAdicionales.length;
        //Subtotales
        var totales = [];
        //Checks
        var idsChecks = [];
        //agregado por diego
        idsChecks = (listacheck != null ? listacheck :[]);
        var filasChecks = [];
        //Ordenacion
        var tipoOrden = 0; //0: ascendente, 1: descendente
        var colOrden = 0; //0: Primera Columna, 1: Segunda Columna
        //Paginacion Simple
        var indicePagina = 0;
        //Paginacion Por Bloques
        var indiceBloque = 0;
        iniciarGrilla();

        function filtrarMatriz() {
            indicePagina = 0;
            indiceBloque = 0;
            crearMatriz();
            mostrarMatriz();
        }

        function iniciarGrilla() {
            crearTabla();
            filtrarMatriz();
        }

        function crearTabla() {
            var html = "";        
            // //contenido de boton Exportar

            html += "<div class='Mensaje' style='display: flex;flex-wrap: nowrap;align-items: center;'>";
            html += "<span>";
            html += mensajeRegistros;
            html += "</span>&nbsp;";
            html += "<span id='spnTotal";
            html += id;
            html += "'></span>";
            if (tieneFiltros) {
                html += "<input type='button' id='btnBorrarFiltro";
                html += id;
                html += "' class='BotonL' value='Borrar Filtros'/>";
            }
            html += "</div>";
            html += "<table class='table'>";
            var cabeceras = lista[0].split("|");
            var anchos = lista[1].split("|");
            tipos = lista[2].split("|");
            nCampos = cabeceras.length;
            html += "<thead class='thead'>";
            html += "<tr class=' ";
            html += id;
            html += "'>";
            if (tieneCheck) {
                html += "<th style='width:30px'>";
                html += "<div>";
                if (tieneCheckCabecera) {
                    html += "<input id='chkCabecera ";
                    html += id;
                    html += "' type='checkbox'/>";
                }               
                html += "</div>";
                html += "</th>";
            }
            for (var j = 0; j < nCampos; j++) {
                html += "<th style='";
               /* html += 'max-widt*/
                html += 'width:';
                html += (+anchos[j]);
                html += "px";
                if (j == 0 && tieneCodigoOculto) {
                    html += ";display:none";
                }
                html += "'>";
                html += "<div>";
                if (tipoGrilla == "EDITABLE") {
                    html += "<div style='justify-content: center;'>";
                } else {
                    html += "<div>";
                }        
                html += "<span class='Enlace ";
                html += id;
                html += "' data-orden='";
                html += j;
                html += "'>";
                html += cabeceras[j];
                html += "</span>";
                /*        html += "&nbs*/
                if (tipoGrilla != "EDITABLE") {
                    html += "<span></span>";
                }             
                html += "</div>";
                if (tieneFiltros) {
                  /*  html += "<br/>";*/
                    if (indices.length > 0 && indices.indexOf(j) > -1) {
                        //if (indiceTipos != null) {

                        //////    switch (indiceTipos[j]) {
                        ////        case 'cbo':
                        //            html += "<select class='Cabecera Combo ";
                        //            html += id;
                        //            html += "'></select>";
                        //            break;
                        //        case 'cbs':
                        //            html += "<select class='Cabecera Combo ";
                        //            html += id;
                        //            html += "'></select>";
                        //            break;
                        //        default:
                        //    }


                     /*   } else {*/
                            html += "<select class='Cabecera Combo ";
                            html += id;
                            html += "'></select>";
                    /*    }*/
                       
                    }
                    else {
                        html += "<input type='text' class='Cabecera Texto ";
                        html += id;
                        html += "'/>";
                    }
                }
                html += "</div>";
                html += "</th>";

            }
            if (nBotones > 0) {
                for (var j = 0; j < nBotones; j++) {
                    html += "<th>";
                    html += botones[j].cabecera;
                    html += "</th>";
                }
            }
            if (tieneNuevoEditar) {
                html += "<th style='width:50px' class='Mostrar'>";
                html += "<img id='btnNuevo";
                html += id;
                html += "' src='";
                html += hdfRaiz.value;
                html += "Images/add.svg' class='Icono Centro NoImprimir ' title='Nuevo'/>";
                html += "</th>";
            }
            if (tieneEliminar) {
                html += "<th style='width:50px'>";
                //html += "<img id='btnEliminar";
                //html += id;
                //html += "' src='";
                //html += hdfRaiz.value;
                //html += "Images/delete.svg' class='Icono Centro NoImprimir ocultar' title='Eliminar Todo'/>";
                html += "</th>";
            }


            if (nlistaBotonesAdicionales > 0) {
                for (var k = 0; k < nlistaBotonesAdicionales; k++) {
                    html += "<th >";
                    html += "</th>";
                }
            }


            html += "</tr>";
            html += "</thead>";
            html += "<tbody class='tbody' id='tbData";
            html += id;
            html += "'>";
            html += "</tbody>";
            html += "<tfoot class='tfoot' style='font-size: 20px;color: var(--color-muniz-primary-verde);'>";
            if (subtotales.length > 0) {
                var nCols = (tieneCheck ? nCampos + 1 : nCampos);
                var n = (tieneCheck ? 1 : 0);
                html += "<tr class='FilaCabecera'>";
                var ccs = 0;
                for (var j = 0; j < nCols; j++) {
                    if (tieneCodigoOculto == true && j == 0) continue;
                    html += "<th class='Derecha'";
                    if (subtotales.indexOf(j - n) > -1) {
                        html += " id='total";
                        html += id;
                        html += subtotales[ccs];
                        html += "'";
                        ccs++;
                    }
                    html += ">";
                    if (subtotales.indexOf(j - n) > -1) {
                        html += "0";
                    }
                    html += "</th>";
                }
                if (tieneNuevoEditar) html += "<th></th>";
                if (tieneEliminar) html += "<th></th>";
                html += "</tr>";
            }
            var nCols = (tieneCheck ? nCampos + 1 : nCampos);
            nCols = (tieneNuevoEditar ? nCols + 1 : nCols);
            nCols = (tieneEliminar ? nCols + 1 : nCols);
            html += "<tr>";
            html += "<td id='tdPagina";
            html += id;
            html += "' colspan='";
            html += nCols;
            html += "' class='Centro'>";
            html += "</td>";
            html += "</tr>";

            html += "</tfoot>";
            html += "</table>";

            //contenido de boton Exportar
            html += "<div class='contenido_buton_export'>";
            if (lista.length > 3 && lista[3] != "") {
                if (tieneExportacion) {
                    //html += "<button id='btnExportarTxt";
                    //html += id;
                    //html += "' class='BotonL'>Exp Txt <img   src='";
                    //html += hdfRaiz.value;
                    //html += "Images/txt.svg' class='Icono Centro NoImprimir' title='Texto'/></button>";
                    html += "<button id='btnExportarXlsx";
                    html += id;
                    html += "' class='BotonL'>Exp Xlsx <img  src='";
                    html += hdfRaiz.value;
                    html += "Images/excel.svg' class='Icono Centro NoImprimir' title='Excel'/></button>";
                    html += "<button id='btnExportarDocx";
                    html += id;
                    html += "' class='BotonL'>Exp Docx <img  src='";
                    html += hdfRaiz.value;
                    html += "Images/word.svg' class='Icono Centro NoImprimir' title='Docs'/></button>";
                    html += "<button id='btnExportarPdf";
                    html += id;
                    html += "' class='BotonL'>Exp Pdf <img  src='";
                    html += hdfRaiz.value;
                    html += "Images/pdf.svg' class='Icono Centro NoImprimir' title='Pdf'/></button>";
                }
                //if (tieneImprimir) {
                //    html += "<button id='btnImprimir";
                //    html += id;
                //    html += "' class='BotonL'>Imprimir <img src='";
                //    html += hdfRaiz.value;
                //    html += "Images/print.svg' class='Icono Centro NoImprimir' title='Imprimir'/></button>";
                //}
            }
            html += "</div>";
             //contenido de boton Exportar

            div.innerHTML = html;
            if (indices.length > 0) llenarCombos();
            if (tipoGrilla!="EDITABLE") configurarOrden();
        }

        function llenarCombos() {
            var combos = document.getElementsByClassName("Cabecera Combo " + id);
            var nCombos = combos.length;
            for (var j = 0; j < nCombos; j++) {

                GUI.Combo(combos[j], ayudas[j], "TODOS");

               

            }
        }

        function llenarCombosSearch() {
            if (indices != null || indices != "") {
                var nCols = indices.length;
                var inicio = indicePagina * registrosPagina;
                for (var i = 0; i < nCols; i++) {
                    var combos = document.getElementsByClassName("InputcboSeacrh " +id+indices[i]);
                    var ncombos = combos.length;
                    for (var j = 0; j < ncombos; j++) {
                        texto = matriz[inicio + j][indices[i]];
                        if (ayudas[i] != null && ayudas[i].length > 0) {
                            GUI.ComboSeacrh(combos[j], ayudas[i], texto);
                            combos[j].value = texto;
                        } 
                        combos[j].setAttribute("data-val", texto);
                        combos[j].setAttribute("data-col", indices[i]);
                        combos[j].setAttribute("data-row", j);
                        combos[j].onchange = function () {
                            cambiarComboRegistro(id, this, this.getAttribute("data-col"), this.getAttribute("data-row"), matriz);
                        }

                    }
                }
                iniciarComboSearchRegistro(id);
            }
        }

        function configurarInputGrilla() {
            for (var i = 0; i < nCampos; i++) {
                var inputs = document.getElementsByClassName("Registro Input " + id + i);
                var nInputs = inputs.length;
                if (nInputs > 0) {
                    for (var j = 0; j < nInputs; j++) {
                        inputs[j].onchange = function () {
                            cambiarInputRegistro(id, this, this.getAttribute("data-col"), this.getAttribute("data-row"), matriz);
                        }
                    }
                  
                }
             }

        }

        function configurarBotonesAdicionales() {
            var cantidadReg = nlistaBotonesAdicionales + nCampos + 2;
            var i = nCampos + 2;
            for (i;i< cantidadReg; i++) {
                var inputs = document.getElementsByClassName("btnBoton" + id + i);
                var nInputs = inputs.length;
                if (nInputs > 0) {
                    for (var j = 0; j < nInputs; j++) {
                        inputs[j].onclick = function () {
                            console.log(id);
                            pulsarBotonGrilla(id, this, this.getAttribute("data-col"), this.getAttribute("data-row"), matriz);
                        }
                    }

                }
            }

        }

        function llenarCombosGrilla() {
            if (indices != null || indices != "") {
                var nCols = indices.length;
                var inicio = indicePagina * registrosPagina;
                var texto = "";
                for (var i = 0; i < nCols; i++) {
                    var combos = document.getElementsByClassName("Registro Combo "+id+indices[i]);
                    var ncombos = combos.length;
                    for (var j = 0; j < ncombos; j++) {
                        texto = matriz[inicio + j][indices[i]];
                        if (ayudas[i] != null && ayudas[i].length>0) {
                            GUI.Combo(combos[j], ayudas[i]);                           
                            combos[j].value = texto;                         
                        } 
                        combos[j].setAttribute("data-val", texto);
                        combos[j].setAttribute("data-col", indices[i]);
                        combos[j].setAttribute("data-row", j);
                        combos[j].onchange = function () {
                            cambiarComboRegistro(id, this, this.getAttribute("data-col"), this.getAttribute("data-row"), matriz);

                        }
                       
                    }
                }
                iniciarComboRegistro(id);
            }
        }

        function buscaValorCombo(combo, texto) {
            var valor = "";
            var n = combo.options.length;
            for (var i = 0; i < n; i++) {
                if (texto == combo.options[i].text) {
                    valor = combo.options[i].value;
                    break;
                }
            }
            return valor;
        }

        function crearMatriz() {
            matriz = [];
            var campos = [];
            var fila = [];
            var esNumero = false;
            var esFecha = false;
            var esTime = false;
            if (subtotales.length > 0) {
                totales = [];
                var nSubtotales = subtotales.length;
                for (var j = 0; j < nSubtotales; j++) {
                    totales.push(0);
                }
            }
            if (lista.length > 3 && lista[3] != "") {
                var cabeceras = document.getElementsByClassName("Cabecera " + id);
                var nCabeceras = cabeceras.length;
                var valores = [];
                for (var j = 0; j < nCabeceras; j++) {
                    if (cabeceras[j].className.indexOf("Texto") > -1) {
                        valores.push(cabeceras[j].value.toLowerCase());
                    }
                    else {
                        valores.push(cabeceras[j].options[cabeceras[j].selectedIndex].value);
                    }
                }
                var exito = false;
                var ccs = 0;
                for (var i = 3; i < nRegistros; i++) {
                    campos = lista[i].split("|");
                    exito = true;
                    for (var j = 0; j < nCabeceras; j++) {
                        if (cabeceras[j].className.indexOf("Texto") > -1) {
                            exito = (valores[j] == "" || campos[j].toString().toLowerCase().indexOf(valores[j]) > -1);
                        }
                        else {
                            exito = (valores[j] == "" || campos[j] == valores[j]);
                        }
                        if (!exito) break;
                    }
                    ccs = 0;
                    if (exito) {
                        fila = [];
                        for (var j = 0; j < nCampos; j++) {
                            esNumero = (tipos[j].indexOf("Int") > -1 || tipos[j].indexOf("Decimal") > -1);
                            esFecha = (tipos[j].indexOf("DateTime") > -1);
                            esTime = (tipos[j].indexOf("Time") > -1);
                            if (esNumero) {
                                fila.push(campos[j] * 1);
                            }
                            else if (esFecha) {                             
                                    fila.push(crearFecha(campos[j]));
                             
                            }
                            else {
                                fila.push(campos[j]);
                            }
                            
                            if (subtotales.length > 0 && subtotales.indexOf(j) > -1) {
                                valor = fila[j];
                                if (esTime) {
                                    totales[ccs] = sumarHorasMinutos(totales[ccs], valor);
                                } else {
                                    totales[ccs] += valor;
                                }
                                ccs++;
                            }
                        }
                        matriz.push(fila);
                    }
                }
            }
        }

        function sumarHorasMinutos(tiempoAcumulado, tiempo) {
            var listaTiempoAcumulado = tiempoAcumulado==0? [0,0]: tiempoAcumulado.split(":");
            var listaTiempo = tiempo == 0 ? [0, 0] : tiempo.split(":");
            var totalMinutos = listaTiempoAcumulado[1] * 1 + listaTiempo[1] * 1;
            var horas = listaTiempoAcumulado[0] * 1 + listaTiempo[0] * 1 + parseInt(totalMinutos/60);
            var minutos = (totalMinutos % 60);
            var time = horas.toString().padStart(2, "0") + ":" + minutos.toString().padStart(2, "0");
            return time;
        }

        function crearFecha(strFecha) {
            if (strFecha == "") return "";
            var backSlash = strFecha.indexOf('/');
            if (backSlash != -1) {

                var fechas = strFecha.split(" ");
                var fecha = fechas[0].split("/");
                var hms=[];
                var ampm="";
                var hora = 0;
                var min = 0;
                var seg = 0;
                if (fechas.length > 1) {
                     hms = fechas[1].split(":");
                     ampm = fechas[2];
                    if (hms.length > 1) {
                         hora = hms[0] * 1;
                        if (ampm == "PM") hora = hora + 12;
                         min = hms[1] * 1;
                         seg = hms[2] * 1;
                    }
                }

                var dia = fecha[0] * 1;
                var mes = +fecha[1] - 1;
                var anio = Number(fecha[2]);

                var fechaDMY = new Date(anio, mes, dia, hora, min, seg);
                return fechaDMY;
            }
            else {

                var fechas = strFecha.split(" ");
                var fecha = fechas[0].split("-");
                var hms;
                var ampm;
                var hora = 0;
                var min = 0;
                var seg = 0;
                if (fechas.length>1) {
                     hms = fechas[1].split(":");
                     ampm = fechas[2];
                    if (hms.length > 1) {
                         hora = hms[0] * 1;
                        if (ampm == "PM") hora = hora + 12;
                         min = hms[1] * 1;
                         seg = hms[2] * 1;
                    }
                }
                

                var anio = Number(fecha[0]);
                var mes = +fecha[1] - 1;
                var dia = fecha[2] * 1;
                var fechaDMY = new Date(anio, mes, dia, hora, min, seg);
                return fechaDMY;
            }

           


        }

        function mostrarMatriz(nuevaFila = false,inicioRecarga) {
            var html = "";
            var nRegMatriz = matriz.length;
            var esNumero = false;
            var esFecha = false;
            var esDecimal = false;
            var existeIdCheck = false;
            var inicio = indicePagina * registrosPagina;
            var fin = inicio + registrosPagina;
            //agregado por diego
            var i;
            
          
            if (nuevaFila) {
                i = nRegMatriz - 1;
            }
            else {
                i = inicio;
            }

            if (inicioRecarga != null) {
                i = inicioRecarga;
            }

            for (i; i < fin; i++) {
                if (i < nRegMatriz) {
                    html += "<tr class='FilaDatos ";
                    html += id;
                    html += "' style='";
                    if (listaColorFila) {
                        html += "background-color:";
                        var nlistaColorFila = listaColorFila.length;
                        var lcolorCamp;
                        for (var x = 0; x < nlistaColorFila; x++) {
                           //lcolorCamp = nlistaColorFila[x].split('|');
                             lcolorCamp = listaColorFila[x];
                            if (matriz[i][lcolorCamp[0]] == lcolorCamp[1]) {
                                html += lcolorCamp[2];
                            }
                        }
                        html += ";";
                    }
                    if (listaColorTextoFila) {
                        html += "color:";
                        var nlistaColorTextoFila = listaColorTextoFila.length;
                        var lcolorCamp;
                        for (var x = 0; x < nlistaColorTextoFila; x++) {
                             // lcolorCamp = listaColorTextoFila[x].split('|');
                           lcolorCamp = listaColorTextoFila[x];
                            if (matriz[i][lcolorCamp[0]] == lcolorCamp[1]) {
                                html += lcolorCamp[2];
                            }
                        }
                        html += ";";
                       
                    }
                    html += "'>";
                    if (tieneCheck) {
                        existeIdCheck = (idsChecks.indexOf(matriz[i][0]) > -1);
                        html += "<td>";
                        html += "<input type='checkbox' class='Check ";
                        html += id;
                        html += "' ";
                        if (existeIdCheck) {
                            html += "checked='checked'";
                        }
                        html += "/>";
                        html += "</td>";
                    }
                    for (var j = 0; j < nCampos; j++) {
                        html += "<td class='";
                        esNumero = (tipos[j].indexOf("Int") > -1 || tipos[j].indexOf("Decimal") > -1);
                        esFecha = (tipos[j].indexOf("DateTime") > -1);
                        esDecimal = (tipos[j].indexOf("Decimal") > -1);
                        esTime = (tipos[j].indexOf("Time") > -1);
                        if (esNumero) {
                            html += "Derecha";
                        }
                        else if (esFecha) {
                            html += "Centro";
                        }
                        else if (esTime) {
                            html += "Centro";
                        }
                        else {
                            html += "Izquierda";
                        }
                        html += "'";
                        if (j == 0 && tieneCodigoOculto) {
                            html += " style='display:none'";
                        }
                        html += ">";

                        switch (tipoGrilla) {

                            case "EDITABLE":
                                if (indices.length > 0 && indices.indexOf(j) > -1) {
                                    if (indiceTipos != null) {

                                        switch (indiceTipos[j]) {
                                            case 'cbo':
                                                html += "<select class='GE Registro Combo ";
                                                html += id;                                              
                                                html += j;
                                                if (!nuevaFila) {
                                                    html +=" colorDisabled "
                                                }
                                                html += "' ";
                                                if (!nuevaFila) {
                                                    html += "";
                                                }
                                                html+="></select > ";
                                                break;
                                            case 'cbs':
                                                html += "<div class='cbosearch";
                                                //html +=  j;
                                                html += "' style='width:100%;'>";
                                                html += "<div class='content_seacrh_button'>";
                                                html += "<div class='GE ";
                                                if (!nuevaFila) {
                                                    html += " colorDisabled "
                                                }
                                                html += "content_button' style='display:flex;'>";
                                                html += "<span type='text' class='GE ";
                                                if (!nuevaFila) {
                                                    html += " colorDisabled "
                                                }
                                                html +="input_text'>";
                                                html += matriz[i][j];
                                                html += "</span> ";
                                                html += "<div class='button_seacrh'>▲</div>";
                                                html += "<input  type='hidden' value=''";
                                                html += " class='GE InputcboSeacrh ";
                                                html += id;
                                                html += j;
                                                html += "' data-campo='";
                                                html += "'  ";
                                                if (!nuevaFila) {
                                                    html += "";
                                                }
                                                html +="/>";
                                                html += "</div>";
                                                html += "<div class='content_seacrh' style='display: none;'>";
                                                html += " <input type='text' class='input_seacrh' placeholder='escriba aqui' />";
                                                html += " <ul class='seacrh_list'>";
                                                html += " </ul>";
                                                html += "</div>";
                                                html += "</div>";
                                                html += "</div>";
                                                break;
                                            default:
                                        }


                                    } else {
                                        html += "<select class='GE ";
                                        if (!nuevaFila) {
                                            html += " colorDisabled "
                                        }
                                        html +=" Cabecera Combo ";
                                        html += id;
                                        html += "'  ";
                                        if (!nuevaFila) {
                                            html += "";
                                        }
                                        html +="></select>";
                                    }

                                }
                                else {                             
                                    switch (indiceTipos[j]) {
                                        case "tme":
                                            html += "<input type='Time' style='width:100%'  class='GE ";
                                            if (!nuevaFila) {
                                                html += " colorDisabled "
                                            }
                                            html +=" Registro Input ";
                                            html += id;
                                            html += j;
                                            html += "' data-row='";
                                            html += i;
                                            html += "' data-col='";
                                            html += j;
                                            html+="' data-val='";
                                            if (esDecimal) html += matriz[i][j].toFixed(2);
                                            else if (esFecha) html += mostrarFechaDMY(matriz[i][j]);
                                            else html += matriz[i][j];
                                            html+="' value='";
                                            if (esDecimal) html += matriz[i][j].toFixed(2);
                                            else if (esFecha) html += mostrarFechaDMY(matriz[i][j]);
                                            else html += matriz[i][j];
                                            html += "'  ";
                                            if (!nuevaFila) {
                                                html += "disabled";
                                            }
                                            html +="/>";
                                            break;
                                        case "lbl":
                                            html += "<label style='width:100%' class='GE Registro Input ";
                                            html += id;
                                            html += i;
                                            html += "' data-row='";
                                            html += i;
                                            html += "' data-col='";
                                            html += j;
                                            html += "' data-val='";
                                            if (esDecimal) html += matriz[i][j].toFixed(2);
                                            else if (esFecha) html += mostrarFechaDMY(matriz[i][j]);
                                            else html += matriz[i][j];
                                            html += "'>";
                                            if (esDecimal) html += matriz[i][j].toFixed(2);
                                            else if (esFecha) html += mostrarFechaDMY(matriz[i][j]);
                                            else html += matriz[i][j];
                                            html += "</label>";
                                        
                                            break;
                                        case "hdn":
                                            html += "<input type='text' style='width:100%' class=' ";
                                            if (!nuevaFila) {
                                                html += " colorDisabled "
                                            }
                                            html += " Registro Input ";
                                            html += id;
                                            html += i;
                                            html += "' data-row='";
                                            html += i;
                                            html += "' data-col='";
                                            html += j;
                                            html += "' data-val='";
                                            if (esDecimal) html += matriz[i][j].toFixed(2);
                                            else if (esFecha) html += mostrarFechaDMY(matriz[i][j]);
                                            else html += matriz[i][j];
                                            html += "' value='";
                                            if (esDecimal) html += matriz[i][j].toFixed(2);
                                            else if (esFecha) html += mostrarFechaDMY(matriz[i][j]);
                                            else html += matriz[i][j];
                                            html += "'  ";
                                            if (!nuevaFila) {
                                                html += "disabled";
                                            }
                                            html += "/>";
                                            break;
                                        default:
                                            html += "<input type='text' style='width:100%' maxlength='";
                                            html += listaMaxlenth[j];
                                            html +="' class='valCharEsp Upper GE ";
                                            if (!nuevaFila) {
                                                html += " colorDisabled "
                                            }
                                            html +=" Registro Input ";
                                            html += id;
                                            html += i;
                                            html += "' data-row='";
                                            html += i;
                                            html += "' data-col='";
                                            html += j;
                                            html += "' data-val='";
                                            if (esDecimal) html += matriz[i][j].toFixed(2);
                                            else if (esFecha) html += mostrarFechaDMY(matriz[i][j]);
                                            else html += html_escape(matriz[i][j]);
                                            html += "' value='";
                                            if (esDecimal) html += matriz[i][j].toFixed(2);
                                            else if (esFecha) html += mostrarFechaDMY(matriz[i][j]);
                                            else html += html_escape(matriz[i][j]);
                                            html += "'  ";
                                            if (!nuevaFila) {
                                                html += "disabled";
                                            }
                                            html +="/>";
                                            break;
                                    }
                                }
                                break;
                            case "NORMAL":
                                if (esDecimal) html += matriz[i][j].toFixed(2);
                                else if (esFecha) html += mostrarFechaDMY(matriz[i][j]);
                                else html += matriz[i][j];
                                break;

                            default:
                        }
                        //grilla editable
                      
                        html += "</td>";
                    }
                    if (nBotones > 0) {
                        for (var j = 0; j < nBotones; j++) {
                            html += "<td>";
                            html += "<button class='BotonGrilla ";
                            html += id;
                            html += "'>";
                            html += botones[j].texto;
                            html += "</button>";
                            html += "</td>";
                        }
                    }
                    if (tieneNuevoEditar) {
                        html += "<td>";
                        html += "<img src='";
                        html += hdfRaiz.value;
                        html += "Images/edit.svg' style='text-align: center;display: inherit;' class='Icono Centro Editar NoImprimir ";
                        html += id;
                        html += "' title='Editar'/>";
                        html += "</td>";
                    }
                    if (tieneEliminar) {
                        html += "<td>";
                        html += "<img src='";
                        html += hdfRaiz.value;
                        html += "Images/delete.svg' class='Icono Centro Eliminar NoImprimir ";
                        html += id;
                        html += "' title='Eliminar'/>";
                        html += "</td>";
                    }
                    if (nlistaBotonesAdicionales > 0) {
                        for (var k = 0; k < nlistaBotonesAdicionales; k++) {
                            html += "<td style='width:50px' class='Mostrar'>";
                            html += "<img class='Icono Centro  btnBoton";
                            html += id + (j + k + 2)*1;
                            html += "' src='";
                            html += hdfRaiz.value;
                            html += listaBotonesAdicionales[k][0];
                            html += "' data-row='";
                            html += i;
                            html += "' data-col='";
                            html += j+k+2;
                            html += "'  title='";
                            html += listaBotonesAdicionales[k][1];
                            html += "'/> ";
                            html += "</td>";
                        }
                    }




                    html += "</tr>";
                }
                else break;
            }


            var tbData = document.getElementById("tbData" + id);
            if (nuevaFila) {
                if (tbData != null) tbData.insertAdjacentHTML('beforeend',html);
            }
            else {
                if (inicioRecarga != null) {
                    if (tbData != null) tbData.insertAdjacentHTML('beforeend', html);
                }
                else {
                    if (tbData != null) tbData.innerHTML = html;
                }
           
            }

            var spTotal = document.getElementById("spnTotal" + id);
            if (spTotal != null) spTotal.innerHTML = matriz.length;
            if (nBotones > 0) configurarBotones();
            if (nlistaBotonesAdicionales > 0) configurarBotonesAdicionales();
            if (subtotales.length > 0) {
                var nSubtotales = subtotales.length;
                var tipo;
                var esDecimal;
                var esTime;
                for (var j = 0; j < nSubtotales; j++) {
                    var celdaSubtotal = document.getElementById("total" + id + subtotales[j]);
                    if (celdaSubtotal != null) {
                        tipo = tipos[subtotales[j]];
                        esDecimal = (tipo.indexOf("Decimal") > -1);
                        esTime = (tipo.indexOf("Time") > -1);
                        if (esDecimal) celdaSubtotal.innerText = totales[j].toFixed(2);
                        else if (esTime) celdaSubtotal.innerText = totales[j];
                        else celdaSubtotal.innerText = totales[j];
                    }
                }
            }
            configurarEventos();
            configurarPaginacion();
            if (tipoGrilla == 'EDITABLE') {
                llenarCombosGrilla();
               
                llenarCombosSearch();
                configurarInputGrilla();
            }
            
        }
    

        function mostrarFechaDMY(fecha) {
            if (fecha == "") return "";
            var anio = fecha.getFullYear();
            var mes = fecha.getMonth() + 1;
            mes = (mes<10 ? "0" + mes.toString() : mes.toString());
            var dia = fecha.getDate();
            dia = (dia < 10 ? "0" + dia.toString() : dia.toString());
            
            var hora = fecha.getHours();
            var strFecha = "";
            if (hora == 0) {
                strFecha = dia + "/" + mes + "/" + anio;
            }
            else {
                var ampm = "AM";
                if (hora > 12) {
                    hora -= 12;
                    ampm = "PM";
                }
                var min = fecha.getMinutes();
                var seg = fecha.getSeconds();
                strFecha = dia + "/" + mes + "/" + anio + " " + hora + ":" + min + ":" + seg + " " + ampm;
            }
          
            return strFecha;
        }

        function configurarBotones() {
            var btns = document.getElementsByClassName("BotonGrilla " + id);
            var nBtns = btns.length;
            for (var j = 0; j < nBtns; j++) {
                btns[j].onclick = function () {
                    var n = (tieneCheck ? 1 : 0);
                    var fila = this.parentNode.parentNode;
                    var idRegistro = fila.childNodes[n].innerText;
                    seleccionarBoton(id, idRegistro, this.innerText);
                }
            }
        }

        function configurarEventos() {
            var filas = document.getElementsByClassName("FilaDatos " + id);
            var nFilas = filas.length;
            for (var i = 0; i < nFilas; i++) {
                filas[i].onclick = function () {
                    var n = (tieneCheck ? 1 : 0);
                    var idRegistro = this.childNodes[n].innerText;
                    if (filaActual != null) {
                        filaActual.className = "FilaDatos " + id;

                    }
                    this.className = "FilaSeleccionada " + id;
                    filaActual = this;
                  seleccionarFila(id, idRegistro, this,n);
                }

               
            }

            var cabeceras = document.getElementsByClassName("Cabecera " + id);
            var nCabeceras = cabeceras.length;
            for (var j = 0; j < nCabeceras; j++) {
                if (cabeceras[j].className.indexOf("Texto") > -1) {
                    cabeceras[j].onkeyup = function (event) {
                        filtrarMatriz();
                        if (tieneCalculofiltroFueraGrilla) {
                            calcularFiltroFueraGrilla(matriz);
                        }
                    }
                }
                else {
                    cabeceras[j].onchange = function (event) {
                        filtrarMatriz();
                        if (tieneCalculofiltroFueraGrilla) {
                            calcularFiltroFueraGrilla(matriz);
                        }
                    }
                }
            }

            var btnBorrarFiltro = document.getElementById("btnBorrarFiltro" + id);
            if (btnBorrarFiltro != null) {
                btnBorrarFiltro.onclick = function () {
                    var cabeceras = document.getElementsByClassName("Cabecera " + id);
                    var nCabeceras = cabeceras.length;
                    for (var j = 0; j < nCabeceras; j++) {
                        cabeceras[j].value = "";
                    }
                    filtrarMatriz();
                    if (tieneCalculofiltroFueraGrilla) {
                        calcularFiltroFueraGrilla(matriz);
                    }
                }
            }

            //var btnExportarTxt = document.getElementById("btnExportarTxt" + id);
            //if (btnExportarTxt != null) {
            //    btnExportarTxt.onclick = function () {
            //        var data = obtenerData("\r\n", ",");
            //        FileSystem.download(data, id + ".txt");
            //    }
            //}

            var btnExportarXlsx = document.getElementById("btnExportarXlsx" + id);
            if (btnExportarXlsx != null) {
                btnExportarXlsx.onclick = function () {
                    var archivo = id + ".xlsx";
                    var data = obtenerData("¬", "|", true);
                    var SubTotalExpoExel = (typeof(tieneSubTotalExportar) == 'undefined' ? "" : tieneSubTotalExportar);
                    var frm = new FormData();
                    frm.append("Data", data);
                    Http.postDownload("Exportacion/Exportar?archivo=" + archivo + "&entrada=" + SubTotalExpoExel , function (rpta) {
                        FileSystem.download(rpta, archivo);
                    }, frm);
                }
            }

            var btnExportarDocx = document.getElementById("btnExportarDocx" + id);
            if (btnExportarDocx != null) {
                btnExportarDocx.onclick = function () {
                    var archivo = id + ".docx";
                    var data = obtenerData("¬", "|", true);
                    var SubTotalExpoDoc = (typeof(tieneSubTotalExportar) == 'undefined' ? "" : tieneSubTotalExportar);
                    var frm = new FormData();
                    frm.append("Data", data);
                    Http.postDownload("Exportacion/Exportar?archivo=" + archivo + "&entrada=" + SubTotalExpoDoc, function (rpta) {
                        FileSystem.download(rpta, archivo);
                    }, frm);
                }
            }

            var btnExportarPdf = document.getElementById("btnExportarPdf" + id);
            if (btnExportarPdf != null) {
                btnExportarPdf.onclick = function () {
                    var archivo = id + ".pdf";
                    var data = obtenerData("¬", "|", true);
                    var frm = new FormData();
                    frm.append("Data", data);
                    Http.postDownload("Exportacion/Exportar?archivo=" + archivo, function (rpta) {
                        FileSystem.download(rpta, archivo);
                    }, frm);
                }
            }

            var btnImprimir = document.getElementById("btnImprimir" + id);
            if (btnImprimir != null) {
                btnImprimir.onclick = function () {
                    var html = "<table style='width:100%'>";
                    var nRegistros = matriz.length;
                    var cabeceras = lista[0].split("|");
                    var nCabeceras = cabeceras.length;
                    var anchos = lista[1].split("|");
                    html += "<thead>";
                    html += "<tr>";
                    for (var j = 0; j < nCabeceras; j++) {
                        html += "<th style='width:";
                        html += anchos[j];
                        html += "px' style='background-color:lightgray;'>";
                        html += cabeceras[j];
                        html += "</th>";
                    }
                    html += "</tr>";
                    html += "</thead>";
                    html += "<tbody>";
                    for (var i = 0; i < nRegistros; i++) {
                        html += "<tr>";
                        for (var j = 0; j < nCabeceras; j++) {
                            html += "<td>";
                            esFecha = (tipos[j].indexOf("DateTime") > -1);
                            if (esFecha) html += mostrarFechaDMY(matriz[i][j]);
                            else html += matriz[i][j];
                            html += "</td>";
                        }
                        html += "</tr>";
                    }
                    html += "</tbody>";
                    html += "</table>";
                    Impresion.imprimirTabla(html);
                }
            }

            var btnNuevo = document.getElementById("btnNuevo" + id);
            if (btnNuevo != null) {
                btnNuevo.onclick = function () {
                    nuevoRegistro(id);
                }
            }

            var btnsEditar = document.getElementsByClassName("Icono Centro Editar " + id);
            if (btnsEditar != null) {
                var nBtnsEditar = btnsEditar.length;
                for (var j = 0; j < nBtnsEditar; j++) {
                    btnsEditar[j].onclick = function () {
                        var fila = this.parentNode.parentNode;
                        var n = 0;
                        if (tieneCheck) n = 1;
                        var cod = fila.childNodes[n].innerText;
                        editarRegistro(id, cod, fila);
                    }
                }
            }

            var btnsEliminar = document.getElementsByClassName("Icono Centro Eliminar " + id);
            if (btnsEliminar != null) {
                var nBtnsEliminar = btnsEliminar.length;
                for (var j = 0; j < nBtnsEliminar; j++) {
                    btnsEliminar[j].onclick = function () {
                        var fila = this.parentNode.parentNode;
                        var n = 0;
                        if (tieneCheck) n = 1;
                        var cod = fila.childNodes[n].innerText;
                        eliminarRegistro(id, cod, fila);
                    }
                }
            }

            if (tieneCheck) {
                var checks = document.getElementsByClassName("Check " + id);
                var nChecks = checks.length;
                var fila;
                var seleccionado;
                var esNumero = (tipos[0].indexOf("Int") > -1 || tipos[0].indexOf("Decimal") > -1);
                var pos;
                for (var i = 0; i < nChecks; i++) {
                    checks[i].onchange = function () {
                        seleccionado = this.checked;
                        fila = this.parentNode.parentNode;
                        cod = fila.childNodes[1].innerText;
                        if (esNumero) cod = +cod;
                        if (seleccionado) {
                            idsChecks.push(cod);
                            filasChecks.push(buscarCodigo(cod));
                        }
                        else {
                            pos = idsChecks.indexOf(cod);
                            if (pos > -1) {
                                idsChecks.splice(pos, 1);
                                filasChecks.splice(pos, 1);
                            }
                        }
                        seleccionarCheck(cod, fila,seleccionado);
                    }
                }

                var chkCabecera = document.getElementById("chkCabecera " + id);
                if (chkCabecera != null) {
                    chkCabecera.onchange = function () {
                        var seleccionado = this.checked;
                        if (!seleccionado) {
                            idsChecks = [];
                            filasChecks = [];
                        }
                        var fila;
                        var cod;
                        var esNumero = (tipos[0].indexOf("Int") > -1 || tipos[0].indexOf("Decimal") > -1);
                        for (var i = 0; i < nChecks; i++) {
                            checks[i].checked = seleccionado;
                            if (seleccionado) {
                                fila = checks[i].parentNode.parentNode;
                                cod = fila.childNodes[1].innerText;
                                if (esNumero) cod = +cod;
                                idsChecks.push(cod);
                                filasChecks.push(buscarCodigo(cod));
                            }
                        }
                    }
                }
            }
        }

        function obtenerData(sepRegistros, sepCampos, tieneCabeceras) {
            var data = "";
            tieneCabeceras = (tieneCabeceras == null ? false : tieneCabeceras);
            var cabeceras = lista[0].split("|");
            var nCabeceras = cabeceras.length;
            var anchos = lista[1].split("|");
            data += cabeceras.join(sepCampos);
            data += sepRegistros;
            if (tieneCabeceras) {
                data += anchos.join(sepCampos);
                data += sepRegistros;
                data += tipos.join(sepCampos);
                data += sepRegistros;
            }
            var nRegistros = matriz.length;
            for (var i = 0; i < nRegistros; i++) {
                //data += matriz[i].join(sepCampos);
                for (var j = 0; j < nCabeceras; j++) {
                    esFecha = (tipos[j].indexOf("DateTime") > -1);
                    if (esFecha) data += mostrarFechaDMY(matriz[i][j]);
                    else data += matriz[i][j];
                    if (j < nCabeceras - 1) data += sepCampos;
                }
                if (i < nRegistros - 1) data += sepRegistros;
            }
            return data;
        }

        function configurarOrden() {
            var enlaces = document.getElementsByClassName("Enlace " + id);
            var nEnlaces = enlaces.length;
            for (var j = 0; j < nEnlaces; j++) {
                enlaces[j].onclick = function () {
                    ordenarColumna(this);
                }
            }
        }

        function ordenarColumna(span) {
            var orden = span.getAttribute("data-orden");
            colOrden = orden * 1;
            /* var spnSimbolo = span.nextSibling.nextSibling;*/
            var spnSimbolo = span.nextSibling;
            var simbolo = spnSimbolo.innerHTML;
            borrarSimbolosOrdenacion();
            if (simbolo == "") {
                tipoOrden = 0;
                spnSimbolo.innerHTML = "▲";

                ///
                matriz.sort(ordenarMatriz);
                ///
                 
            }
            else {
                if (simbolo == "▲") {
                    tipoOrden = 1;
                    spnSimbolo.innerHTML = "▼";
                }
                else {
                    tipoOrden = 0;
                    spnSimbolo.innerHTML = "▲";
                }
                matriz.reverse();
            }
            mostrarMatriz();
        }

        function borrarSimbolosOrdenacion() {
            var enlaces = document.getElementsByClassName("Enlace " + id);
            var nEnlaces = enlaces.length;
            for (var j = 0; j < nEnlaces; j++) {
               /* enlaces[j].nextSibling.nextSibling.innerHTML = "";*/
                enlaces[j].nextSibling.innerHTML = "";
            }
        }

        function configurarPaginacion() {
            var nRegistros = matriz.length;
            var totalPaginas = Math.floor(nRegistros / registrosPagina);
            if (nRegistros % registrosPagina > 0) totalPaginas++;
            var html = "";
            if (totalPaginas > 1) {
                var totalRegistros = matriz.length;
                var registrosBloque = registrosPagina * paginasBloque;
                var totalBloques = Math.floor(totalRegistros / registrosBloque);
                if (totalRegistros % registrosBloque > 0) totalBloques++;
                if (indiceBloque > 0) {
                    html += "<button class='Pag ";
                    html += id;
                    html += " Pagina' data-pag='-1'>";
                    html += "<<";
                    html += "</button>";
                    html += "<button class='Pag ";
                    html += id;
                    html += " Pagina' data-pag='-2'>";
                    html += "<";
                    html += "</button>";
                }
                var inicio = indiceBloque * paginasBloque;
                var fin = inicio + paginasBloque;
                for (var j = inicio; j < fin; j++) {
                    if (j < totalPaginas) {
                        html += "<button class='Pag ";
                        html += id;
                        html += " ";
                        if (indicePagina == j) html += "PaginaSeleccionada";
                        else html += "Pagina";
                        html += "' data-pag='";
                        html += j;
                        html += "'>";
                        html += (j + 1);
                        html += "</button>";
                    }
                    else break;
                }
                if (indiceBloque < (totalBloques-1)) {
                    html += "<button class='Pag ";
                    html += id;
                    html += " Pagina' data-pag='-3'>";
                    html += ">";
                    html += "</button>";
                    html += "<button class='Pag ";
                    html += id;
                    html += " Pagina' data-pag='-4'>";
                    html += ">>";
                    html += "</button>";
                }
                html += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";
                html += "<select id='cboPagina";
                html += id;
                html += "'>";
                for (var j = 0; j < totalPaginas; j++) {
                    html += "<option value='";
                    html += j;
                    html += "'";
                    if (indicePagina == j) html += " selected";
                    html += ">";
                    html += (j + 1);
                    html += "</option>";
                }
                html += "</select>";
            }
            var tdPagina = document.getElementById("tdPagina" + id);
            if (tdPagina != null) {
                tdPagina.innerHTML = html;
                configurarEventosPagina();
            }
        }

        function configurarEventosPagina() {
            var paginas = document.getElementsByClassName("Pag " + id);
            var nPaginas = paginas.length;
            for (var j = 0; j < nPaginas; j++) {
                paginas[j].onclick = function () {
                    paginar(+this.getAttribute("data-pag"));
                }
            }

            var cboPagina = document.getElementById("cboPagina" + id);
            if (cboPagina != null) {
                cboPagina.onchange = function () {
                    paginar(this.value);
                }
            }
        }

        function paginar(indice) {
            if (indice > -1) {
                indicePagina = indice;
                indiceBloque = Math.floor(indicePagina / paginasBloque);
            }
            else {
                var totalRegistros = matriz.length;
                var registrosBloque = registrosPagina * paginasBloque;
                var totalBloques = Math.floor(totalRegistros / registrosBloque);
                if (totalRegistros % registrosBloque > 0) totalBloques++;
                switch (indice) {
                    case -1: //Primer Bloque
                        indiceBloque = 0;
                        indicePagina = 0;
                        break;
                    case -2: //Bloque Anterior
                        indiceBloque--;
                        indicePagina = indiceBloque * paginasBloque;
                        break;
                    case -3: //Bloque Siguiente
                        indiceBloque++;
                        indicePagina = indiceBloque * paginasBloque;
                        break;
                    case -4: //Ultimo Bloque
                        indiceBloque = (totalBloques - 1);
                        indicePagina = indiceBloque * paginasBloque;
                        break;
                }
            }
            mostrarMatriz();
        }

        this.ObtenerMatriz = function() {
            return matriz;
        }

        this.ObtenerCheckIds = function () {
            var esNumero = (tipos[0].indexOf("Int") > -1 || tipos[0].indexOf("Decimal") > -1);
            if (esNumero) idsChecks.sort(ordenarVector);
            else idsChecks.sort();
            return idsChecks;
        }

        this.ObtenerCheckFilas = function () {
            var esNumero = (tipos[0].indexOf("Int") > -1 || tipos[0].indexOf("Decimal") > -1);
            if (esNumero) filasChecks.sort(ordenarMatrizAscCod);
            else filasChecks.sort();
            var data = [];
            var nFilasChecks = filasChecks.length;
            for (var i = 0; i < nFilasChecks; i++) {
                data.push(filasChecks[i].join("|"));
            }
            return data.join("¬");
        }

        function ordenarVector(x, y) {
            return (x > y ? 1 : -1);
        }

        function ordenarMatrizAscCod(x, y) {
            var idX = x[0];
            var idY = y[0];
            return (idX > idY ? 1 : -1);
        }

        function ordenarMatriz(x, y) {
            var rpta = 0;
            var idX = x[colOrden];
            var idY = y[colOrden];
            if (tipoOrden == 0) rpta = (idX > idY ? 1 : -1);
            else rpta = (idX < idY ? 1 : -1);
            return rpta;
        }

        function buscarCodigo(id) {
            var fila = [];
            var nRegistros = matriz.length;
            if (nRegistros > 0) {
                var nCampos = matriz[0].length;
                for (var i = 0; i < nRegistros; i++) {
                    if (matriz[i][0] == id) {
                        for (var j = 0; j < nCampos; j++) {
                            fila.push(matriz[i][j])
                        }
                        break;
                    }
                }
            }
            return fila;
        }

        this.CambiarValor = function (indiceColBuscar, valorColBuscar, indiceColCambiar, valorColCambiar) {
            var nRegistros = lista.length;
            if (nRegistros > 0) {
                var exito = false;
                var campos = [];
                for (var i = 3; i < nRegistros; i++) {
                    campos = lista[i].split("|");
                    if (campos[indiceColBuscar] == valorColBuscar) {
                        campos[indiceColCambiar] = valorColCambiar;
                        lista[i] = campos.join("|");
                        exito = true;
                        break;
                    }
                }
                if (exito) {
                    crearMatriz();
                    mostrarMatriz();
                }
            }
        }
        this.cargarMatriz = function () {
            mostrarMatriz();
        }

        this.AgregarNuevo = function () {
            mostrarMatriz(true);
        }

        this.vaciarGrilla = function () {
            matriz = [];
            mostrarMatriz();
        }
        this.obtenerFilaSeleccionada = function () {
            return filaActual;
        }

        this.llenarDataMatrizDesdePos = function (pos){
            mostrarMatriz(true, pos);
        }

        this.calcularTotales = function () {
            if (subtotales.length > 0) {
                var nRegMatriz = matriz.length;
                var nSubtotales = subtotales.length;
                var tipo;
                var esDecimal;
                var esTime;
                
                totales = [];
                var ccs = 0;
                for (var j = 0; j < nSubtotales; j++) {
                    totales.push(0);
                }
                for (var i = 0; i < nRegMatriz; i++) {
                    ccs = 0;
                    for (var j = 0; j < nCampos; j++) {
                        if (subtotales.length > 0 && subtotales.indexOf(j) > -1) { 
                           esTime = (tipos[j].indexOf("Time") > -1);
                            if (esTime) {
                                totales[ccs] = sumarHorasMinutos(totales[ccs], matriz[i][j]);
                            } else {
                                totales[ccs] += matriz[i][j];
                            }
                            ccs++;
                        }
                    }
                }        
                for (var j = 0; j < nSubtotales; j++) {
                    var celdaSubtotal = document.getElementById("total" + id + subtotales[j]);
                    if (celdaSubtotal != null) {
                        tipo = tipos[subtotales[j]];
                        esDecimal = (tipo.indexOf("Decimal") > -1);
                        esTime = (tipo.indexOf("Time") > -1);
                        if (esDecimal) celdaSubtotal.innerText = totales[j].toFixed(2);
                        else if (esTime) celdaSubtotal.innerText = totales[j];
                        else celdaSubtotal.innerText = totales[j];
                    }
                }
            }
        }
    }



    GUI.TextList = function (div, lista, idTextList, ancho) {
        var idText = idTextList;
        var ancho = (ancho == null ? 300 : ancho);
        var html = "<input id='txtBusqueda";
        this.IdLista = "";

        html += idTextList;
        html += "' style='width:";
        html += ancho;
        html += "px'/>";
        html += "<div style='display:none;background-Color:white;overflow-y:auto;height:200px; width:";
        html += ancho;
        html += "px'>";
        html += "<ul id='ulTextList";
        html += idTextList;
        html += "'>";
        html += "</ul>";
        html += "</div>";
        div.innerHTML = html;

        var txtBusqueda = document.getElementById("txtBusqueda" + idTextList);
        if (txtBusqueda != null) {
            txtBusqueda.onkeyup = function (event) {
                var ulTextList = document.getElementById("ulTextList" + idTextList);
                ulTextList.parentNode.style.display = "block";
                var nRegistros = lista.length;
                var campos = [];
                var valor = txtBusqueda.value.toLowerCase();
                var data = "";
                for (var i = 0; i < nRegistros; i++) {
                    campos = lista[i].split("|");
                    if (valor == "" || (valor != "" && campos[1].toLowerCase().startsWith(valor))) {
                        data += "<li data-id='";
                        data += campos[0];
                        data += "' class='";
                        data += idTextList;
                        data += " lista' style='cursor:pointer'>";
                        data += campos[1];
                        data += "</li>";
                    }
                }

                if (ulTextList != null) {
                    ulTextList.innerHTML = data;
                    var lis = document.getElementsByClassName(idTextList + " lista");
                    var nlis = lis.length;
                    for (var i = 0; i < nlis; i++) {
                        lis[i].onclick = function () {
                            txtBusqueda.value = this.innerText;
                            txtBusqueda.setAttribute("data-id", this.getAttribute("data-id"));
                            ulTextList.parentNode.style.display = "none";
                        }
                    }
                }
            }
        }
        function obtenerIdBusqueda() {
            var txtBusqueda = document.getElementById("txtBusqueda" + idText);
            if (txtBusqueda != null) {
                return txtBusqueda.getAttribute("data-id");
            }
        }

        this.IdLista = obtenerIdBusqueda;
    }

    GUI.ObtenerDatos = function (claseGrab) {
        var data = "";
        if (claseGrab == null) claseGrab = "G";
        var controles = document.getElementsByClassName(claseGrab);
        var nControles = controles.length;

        for (var i = 0; i < nControles; i++) {
            data += controles[i].value;
            if (i < nControles - 1) data += "|";
        }
        return (data);
    }

    GUI.ObtenerCamposValores = function (claseGrab) {
        var data = "";
        if (claseGrab == null) claseGrab = "G";
        var controles = document.getElementsByClassName(claseGrab);
        var nControles = controles.length;
        var campos = "";
        var valores = "";
        for (var i = 0; i < nControles; i++) {
            campos += controles[i].getAttribute("data-campo");
            //valores += controles[i].value;

            var tipocontrol = controles[i].id.substr(0, 3);
            switch (tipocontrol) {
                case "div":
                    valores += controles[i].innerHTML;
                    break;
                case "tme":
                    var listaTime = tmeMinutos.value.split(":");
                    valores += parseInt(listaTime[0]) * 60 + parseInt(listaTime[1]);
                    break;
                default:
                    valores += controles[i].value;
            }


            if (i < nControles - 1) {
                campos += "|";
                valores += "|";
            }
        }
        data = campos + "~" + valores;
        return (data);
    }

    GUI.LimpiarDatos = function (claseBorrar) {
        if (claseBorrar == null) claseBorrar = "R";
        var controles = document.getElementsByClassName(claseBorrar);
        var nControles = controles.length;
        for (var i = 0; i < nControles; i++) {
            controles[i].value = "";
            controles[i].style.borderColor = "";
            var tipocontrol = controles[i].id.substr(0, 3);
            switch (tipocontrol) {
                case "cbs":
                    var texto = controles[i].parentNode.children[0];
                    texto.innerHTML = "SELECCIONE";
                    break;

                case "div":
                    controles[i].innerHTML="";
                    break;
                case "lbl":
                    controles[i].innerHTML = "";
                    break;
                default:
            }
        }
    }

    GUI.MostrarDatos = function (claseMostrar, valoresMostrar) {
        if (claseMostrar == null) claseMostrar = "G";
        var controles = document.getElementsByClassName(claseMostrar);
        var nControles = controles.length;
        for (var i = 0; i < nControles; i++) {
            if (i < valoresMostrar.length) {
               
                var tipocontrol = controles[i].id.substr(0, 3);
                switch (tipocontrol) {
                    case "cbs":
                        var listaElementos = controles[i].parentNode.parentNode.children[1].children[1];
                        var texto = controles[i].parentNode.children[0];
                        var textoElemento = listaElementos.querySelectorAll('[data-item=' + '"' + valoresMostrar[i] + '"' + ']');
                        texto.innerHTML = textoElemento[0].innerHTML;
                        controles[i].value = valoresMostrar[i];
                        break;
                    case "div":
                        controles[i].innerHTML = valoresMostrar[i];
                        break;
                    default:
                        controles[i].value = valoresMostrar[i];
                }
            }
        }
    }

    GUI.Tree2 = function (div, listaCabecera, listaDetalle, id, tieneNuevoEditar, detallecodigoOculto) {
        var nRegCabeceras = listaCabecera.length;
        var nRegDetalles = listaDetalle.length;
        var cabeceraCampos = listaCabecera[0].split("|");
        var cabeceraAnchos = listaCabecera[1].split("|");
        var nCaberaCampos = cabeceraCampos.length;
        var matrizDetalle = [];
        //agregado tieneNuevoEditar
        tieneNuevoEditar = (tieneNuevoEditar == null ? false : tieneNuevoEditar);
      
        //agregado
        var campos = [];
        var pos = 3;

        var html = "";
        html += "<table style='width:100%' class='table'>";
        html += "<thead class='thead'>";
        html += "<tr class='FilaCabecera'>";
        for (var j = 0; j < nCaberaCampos; j++) {
            html += "<th>";

            html += "<div>";
            html += "<div>";
            if (j == 0) {
                html += "<span id='spnCab' class='CabExpandir NoImprimir' class='Centrado'>+</span>";
            }
            html += "<span class='titlethead'>";
            html += cabeceraCampos[j];;
            html += "</span><br/>";
            html += "<input type='text' class='CabTree";
            html += id;
            html += "'/>";
            html += "</div>";
            html += "</div>";


            //if (j == 0) {
            //    html += "<div>";
            //    html += "<span id='spnCab' class='CabExpandir NoImprimir' class='Centrado'>+</span>";
            //    html += "<span>";
            //    html += cabeceraCampos[j];
            //    html += "</span>";
            //    html += "</div>";
            //}
           /* else html += cabeceraCampos[j];*/
            html += "</th>";
        }
        //agregado 
        if (tieneNuevoEditar) {
            html += "<th style='width:50px'>";
            html += "<img src='";
            html += hdfRaiz.value;
            html += "Images/add.svg' class='Icono treviewAdd";
            html += id;
            html += "' title='Editar'/>";
            html += "</th>";
        }
        //agregado 
        html += "</tr>";
        html += "</thead>";
        html += "<tbody class='tbody' id='tbData";
        html += id;
        html +="'>";
        html += "</tbody>";
        html += "</table>";
        div.innerHTML = html;
        configurarFiltrosTree2();
        crearTablaCabecera();

        function crearTablaCabecera() {
            var html = "";
            var valores = [];
            var cabeceras = document.getElementsByClassName("CabTree" + id);
            var nCabeceras = cabeceras.length;
            for (var i = 0; i < nCabeceras; i++) {
                valores.push(cabeceras[i].value.toLowerCase());
            }
            var exito = false;
            for (var i = 3; i < nRegCabeceras; i++) {
                campos = listaCabecera[i].split("|");
                for (var j = 0; j < nCabeceras; j++) {
                    exito = (valores == "" || campos[j].toLowerCase().indexOf(valores[j]) > -1);
                    if (!exito) break;
                }
                if (exito) {
                    html += "<tr class='FilaDatos'>";
                    for (var j = 0; j < nCaberaCampos; j++) {
                        html += "<td>";
                        if (j == 0) {
                            html += "<div class='form_basic_horizontal'>"
                            html += "<span class='Expandible Centrado NoImprimir' data-id='";
                            html += campos[0];
                            html += "'>+</span>";
                            html += "<span>";
                            html += campos[j];
                            html += "</span>";
                            html += "</div>";
                        }
                        else html += campos[j];
                        html += "</td>";
                    }
                    //agregado 
                    if (tieneNuevoEditar) {
                        html += "<td>";
                        html += "<img src='";
                        html += hdfRaiz.value;
                        html += "Images/edit.svg' class='Icono treviewEdit";
                        html += id;
                        html += "' title='Editar'/>";
                        html += "</td>";
                    }
                    //agregado 
                    html += "</tr>";
                    html += "<tr style='display:none' id='tr";
                    html += campos[0];
                    html += "'>";
                    html += "<td>&nbsp;</td>";
                    html += "<td colspan='";
                    html += nCaberaCampos - 1;
                    html += "'>";
                    html += crearTablaDetalle(campos[0]);
                    html += "</td>";
                    html += "</tr>";
                }
            }
            var tbl = document.getElementById("tbData" + id);
            if (tbl!=null) tbl.innerHTML = html;
        }

        configurarExpandirColapsar();

        var btnsTrewViewAdd = document.getElementsByClassName("treviewAdd" + id);
        if (btnsTrewViewAdd != null) {
            var nbtnsTrewViewAdd= btnsTrewViewAdd.length;
            for (var j = 0; j < nbtnsTrewViewAdd; j++) {
                btnsTrewViewAdd[j].onclick = function () {
                  
                    AgregarTrewViewRegistro();
                }
            }
        }

        var treviewEdit = document.getElementsByClassName("treviewEdit" + id);
        if (treviewEdit != null) {
            var nbtnsTrewViewAdd = treviewEdit.length;
            for (var j = 0; j < nbtnsTrewViewAdd; j++) {
                treviewEdit[j].onclick = function () {
                    var n = 0;
                    var fila = this.parentNode.parentNode;
                    var cod = fila.childNodes[n].innerText.slice(1).trim();
                    editarTrewViewRegistro(cod, this, matrizDetalle[cod]);
                }
            }
        }


        function crearTablaDetalle(idCabecera) {
            //agregado por diego
            matrizDetalle[idCabecera] = [];
            //
            var html = "";
            var detalleCampos = listaDetalle[0].split("|");
            var detalleAnchos = listaDetalle[1].split("|");
            var detalleTipos = listaDetalle[2].split("|");
            var nDetalleCampos = detalleCampos.length - 1;
            var posCampo = nDetalleCampos;
            html += "<table style='width:100%'>";
            html += "<tr class='FilaCabecera'>";
            for (var j = 0; j < nDetalleCampos; j++) {
                html += "<th style='width:";
                html += detalleAnchos[j];
                html += "px;";
                if (j == 0 && detallecodigoOculto) {
                    html += ";display:none";
                }
                html +="'>";
                html += detalleCampos[j];
                html += "</th>";
            }
            html += "</tr>";
            html += "<tbody class='tbody'>";
            var campos = [];
            var tipoCabecera = "";
            for (var i = pos; i < nRegDetalles; i++) {
                campos = listaDetalle[i].split("|");
                

                tipoCabecera = campos[posCampo];
                if (tipoCabecera == idCabecera) {
                    //agregado por diego
                    matrizDetalle[idCabecera].push(listaDetalle[i]);
                    //
                    html += "<tr class='FilaDatos'>";
                    for (var j = 0; j < nDetalleCampos; j++) {
                        tipo = detalleTipos[j];
                        html += "<td ";
                        if (j == 0 && detallecodigoOculto) {
                            html += "style='display:none'";
                        }
                        html += ">";

                        html += campos[j];
                        html += "</td>";
                    }
                    html += "</tr>";
                }
                //else {
                //    pos = i;
                //    break;
                //}
            }
            html += "</tbody>";
            html += "</table>";
            return html;
        }

        function configurarExpandirColapsar() {
            var spans = document.getElementsByClassName("Expandible");
            var nSpans = spans.length;
            for (var i = 0; i < nSpans; i++) {
                spans[i].onclick = function () {
                    var id = this.getAttribute("data-id");
                    var tr = document.getElementById("tr" + id);
                    if (this.innerText == "+") {
                        tr.style.display = "table-row";
                        this.innerText = "-";
                    }
                    else {
                        tr.style.display = "none";
                        this.innerText = "+";
                    }
                }
            }

            var spnCab = document.getElementById("spnCab");
            if (spnCab != null) {
                spnCab.onclick = function () {
                    for (var i = 0; i < nSpans; i++) {
                        var id = spans[i].getAttribute("data-id");
                        var tr = document.getElementById("tr" + id);
                        if (this.innerText == "+") {
                            spans[i].innerText = "-";
                            tr.style.display = "table-row";
                        }
                        else {
                            spans[i].innerText = "+";
                            tr.style.display = "none";
                        }
                    }
                    if (this.innerText == "+") {
                        this.innerText = "-";
                    }
                    else {
                        this.innerText = "+";
                    }
                }
            }
        }

        function configurarFiltrosTree2() {
            var cabeceras = document.getElementsByClassName("CabTree" + id);
            var nCabeceras = cabeceras.length;
            for (var i = 0; i < nCabeceras; i++) {
                cabeceras[i].onkeyup = function () {
                    crearTablaCabecera();
                }
            }
        }
    }

    GUI.Tree3 = function (div, listaCabecera, listaDetalle, listaSubdetalle) {
        var nCabeceras = listaCabecera.length;
        var nDetalles = listaDetalle.length;
        var nSubdetalles = listaSubdetalle.length;

        var cabeceraCampos = listaCabecera[0].split("|");
        var cabeceraAnchos = listaCabecera[1].split("|");
        var nCaberaCampos = cabeceraCampos.length;
        var campos = [];
        var posDetalle = 3;
        var posSubdetalle = 3;

        var html = "";
        html += "<table style='width:100%'>";
        html += "<tr class='FilaCabecera'>";
        for (var j = 0; j < nCaberaCampos; j++) {
            html += "<th style='width:";
            html += cabeceraAnchos[j];
            html += "px'>";
            if (j == 0) {
                html += "<div>";
                html += "<span id='spnCab' class='Expandible NoImprimir' class='Centrado'>+</span>";
                html += "<span>";
                html += cabeceraCampos[j];
                html += "</span>";
                html += "</div>";
            }
            else html += cabeceraCampos[j];
            html += "</th>";
        }
        html += "</tr>";
        for (var i = 3; i < nCabeceras; i++) {
            campos = listaCabecera[i].split("|");
            html += "<tr class='FilaDatos'>";
            for (var j = 0; j < nCaberaCampos; j++) {
                html += "<td>";
                if (j == 0) {
                    html += "<div>";
                    html += "<span class='FilaExpandir Expandible Centrado NoImprimir' data-id='";
                    html += campos[0];
                    html += "'>+</span>";
                    html += "<span>";
                    html += campos[j];
                    html += "</span>";
                    html += "</div>";
                }
                else html += campos[j];
                html += "</td>";
            }
            html += "</tr>";
            html += "<tr style='display:none' id='tr";
            html += campos[0];
            html += "'>";
            html += "<td>&nbsp;</td>";
            html += "<td colspan='";
            html += nCaberaCampos - 1;
            html += "'>";
            html += crearTablaDetalle(campos[0]);
            html += "</td>";
            html += "</tr>";
        }
        html += "</table>";
        div.innerHTML = html;
        configurarExpandirColapsar();

        function crearTablaDetalle(idCabecera) {
            var html = "";
            var detalleCampos = listaDetalle[0].split("|");
            var detalleAnchos = listaDetalle[1].split("|");
            var detalleTipos = listaDetalle[2].split("|");
            var nDetalleCampos = detalleCampos.length - 1;
            var posCampo = nDetalleCampos;
            html += "<table style='width:100%'>";
            html += "<tr class='FilaCabecera'>";
            for (var j = 0; j < nDetalleCampos; j++) {
                html += "<th style='width:";
                html += detalleAnchos[j];
                html += "px'>";
                if (j == 0) {
                    html += "<div>";
                    html += "<span id='spnCab";
                    html += idCabecera;
                    html += "' class='CabExpandir Expandible NoImprimir Centrado' data-id='";
                    html += idCabecera;
                    html += "'>+</span>";
                    html += "<span>";
                    html += detalleCampos[j];
                    html += "</span>";
                    html += "</div>";
                }
                else html += detalleCampos[j];
                html += "</th>";
            }
            html += "</tr>";
            var campos = [];
            var tipo = "";
            for (var i = posDetalle; i < nDetalles; i++) {
                campos = listaDetalle[i].split("|");
                if (campos[posCampo] == idCabecera) {
                    html += "<tr class='FilaDatos'>";
                    for (var j = 0; j < nDetalleCampos; j++) {
                        tipo = detalleTipos[j];
                        html += "<td";
                        if (tipo.startsWith("Int") || tipo.startsWith("Decimal")) {
                            html += " class='Derecha'";
                        }
                        html += ">";
                        if (j == 0) {
                            html += "<div>";
                            html += "<span class='Expandible Centrado NoImprimir FilaExpandir";
                            html += idCabecera;
                            html += "' data-id='";
                            html += campos[0];
                            html += "'>+</span>";
                            html += "<span>";
                            html += campos[j];
                            html += "</span>";
                            html += "</div>";
                        }
                        else html += campos[j];
                        html += "</td>";
                    }
                    html += "</tr>";
                    html += "<tr style='display:none' id='tr";
                    html += campos[0];
                    html += "'>";
                    html += "<td>&nbsp;</td>";
                    html += "<td colspan='";
                    html += nDetalleCampos - 1;
                    html += "'>";
                    html += crearTablaSubdetalle(campos[0]);
                    html += "</td>";
                    html += "</tr>";
                }
                else {
                    posDetalle = i;
                    break;
                }
            }
            html += "</table>";
            return html;
        }

        function crearTablaSubdetalle(idDetalle) {
            var html = "";
            var subdetalleCampos = listaSubdetalle[0].split("|");
            var subdetalleAnchos = listaSubdetalle[1].split("|");
            var subdetalleTipos = listaSubdetalle[2].split("|");
            var nSubdetalleCampos = subdetalleCampos.length - 1;
            var posCampo = nSubdetalleCampos;
            html += "<table style='width:100%'>";
            html += "<tr class='FilaCabecera'>";
            for (var j = 0; j < nSubdetalleCampos; j++) {
                html += "<th style='width:";
                html += subdetalleAnchos[j];
                html += "px'>";
                html += subdetalleCampos[j];
                html += "</th>";
            }
            html += "</tr>";
            var campos = [];
            var tipo = "";
            for (var i = posSubdetalle; i < nSubdetalles; i++) {
                campos = listaSubdetalle[i].split("|");
                if (campos[posCampo] == idDetalle) {
                    html += "<tr class='FilaDatos'>";
                    for (var j = 0; j < nSubdetalleCampos; j++) {
                        tipo = subdetalleTipos[j];
                        html += "<td";
                        if (tipo.startsWith("Int") || tipo.startsWith("Decimal")) {
                            html += " class='Derecha'";
                        }
                        html += ">";
                        html += campos[j];
                        html += "</td>";
                    }
                    html += "</tr>";
                }
                else {
                    posSubdetalle = i;
                    break;
                }
            }
            html += "</table>";
            return html;
        }

        function expandirFilas(clase) {
            var spans = document.getElementsByClassName(clase);
            var nSpans = spans.length;
            for (var i = 0; i < nSpans; i++) {
                spans[i].onclick = function () {
                    var id = this.getAttribute("data-id");
                    var tr = document.getElementById("tr" + id);
                    if (this.innerText == "+") {
                        tr.style.display = "table-row";
                        this.innerText = "-";
                    }
                    else {
                        tr.style.display = "none";
                        this.innerText = "+";
                    }
                }
            }
        }

        function expandirCabecera(clase, idCabecera) {
            var spans = document.getElementsByClassName(clase);
            var nSpans = spans.length;
            var spnCab = document.getElementById(idCabecera);
            if (spnCab != null) {
                spnCab.onclick = function () {
                    for (var i = 0; i < nSpans; i++) {
                        var id = spans[i].getAttribute("data-id");
                        var tr = document.getElementById("tr" + id);
                        if (this.innerText == "+") {
                            spans[i].innerText = "-";
                            tr.style.display = "table-row";
                        }
                        else {
                            spans[i].innerText = "+";
                            tr.style.display = "none";
                        }
                    }
                    if (this.innerText == "+") {
                        this.innerText = "-";
                    }
                    else {
                        this.innerText = "+";
                    }
                }
            }
        }

        function configurarExpandirColapsar() {
            expandirCabecera("FilaExpandir", "spnCab");
            expandirFilas("FilaExpandir");

            var cabs = document.getElementsByClassName("CabExpandir");
            var ncabs = cabs.length;
            for (var i = 0; i < ncabs; i++) {
                var id = cabs[i].getAttribute("data-id");
                expandirCabecera("FilaExpandir" + id, cabs[i].id);
                expandirFilas("FilaExpandir" + id);
            }
        }
    }

    GUI.TreeView = function(div, lista, raiz, check) {
        check = (check == null ? false : check);
        var html = "<ul id='ulMenu' style='cursor:pointer'>";
        html += "<li>";
        if (check) html += "<input type='checkbox' class='check'/>";
        html += raiz;
        html += "<ul>";
        var nregistros = lista.length;
        var campos = [];
        for (var i = 0; i < nregistros; i++) {
            campos = lista[i].split("|");
            if (campos[3] == "0") {
                html += "<li data-id='";
                html += campos[0];
                html += "' data-accion='";
                html += campos[2];
                html += "'>";
                if (check) html += "<input type='checkbox' class='check'/>";
                html += campos[1];
                html += crearSubmenu(lista, campos[0], check);
                html += "</li>";
            }
        }
        html += "</ul>";
        html += "</li>";
        html += "</ul>";
        div.innerHTML = html;
        configurarTreeViewChecks(check);

        function crearSubmenu(lista, idPadre, check) {
            check = (check == null ? false : check);
            var html = "<ul>";
            var nregistros = lista.length;
            var campos = [];
            for (var i = 0; i < nregistros; i++) {
                campos = lista[i].split("|");
                if (campos[3] == idPadre) {
                    html += "<li data-id='";
                    html += campos[0];
                    html += "' data-accion='";
                    html += campos[2];
                    html += "'>";
                    if (check) html += "<input type='checkbox' class='check'/>";
                    html += campos[1];
                    html += crearSubmenu(lista, campos[0], check);
                    html += "</li>";
                }
            }
            html += "</ul>";
            return html;
        }

        function configurarTreeViewChecks(check) {
            var ulMenu = document.getElementById("ulMenu");
            ulMenu.onclick = function (event) {
                var liMenu = event.target;
                if (liMenu.childNodes.length > 1) {
                    var ulSubMenu = liMenu.childNodes[check ? 2 : 1];
                    if (ulSubMenu.hasChildNodes()) {
                        if (ulSubMenu.style.display == "" || ulSubMenu.style.display == "block") {
                            ulSubMenu.style.display = "none";
                        }
                        else {
                            ulSubMenu.style.display = "block";
                        }
                    }
                    else {
                        var opcion = liMenu.innerText;
                        var accion = liMenu.getAttribute("data-accion");
                        seleccionarNodoTreeView(liMenu, opcion, accion);
                    }
                }
            }

            var checks = document.getElementsByClassName("check");
            var nchecks = checks.length;
            for (var i = 0; i < nchecks; i++) {
                var check = checks[i];
                check.onclick = function () {
                    seleccionarCheckTreeView(this);
                }
            }
        }

        function seleccionarCheckTreeView(chkPadre) {
            var seleccion = chkPadre.checked;
            var liPadre = chkPadre.parentNode;
            var ulPadre = liPadre.childNodes[2];
            if (ulPadre.hasChildNodes()) {
                var nHijos = ulPadre.childNodes.length;
                var liHijos = ulPadre.childNodes;
                var liHijo;
                var chkHijo;
                for (var i = 0; i < nHijos; i++) {
                    liHijo = liHijos[i];
                    chkHijo = liHijo.firstChild;
                    chkHijo.checked = seleccion;
                    seleccionarCheckTreeView(chkHijo);
                }
            }
        }

        function obtenerChecksSeleccionados() {
            var checks = document.getElementsByClassName("check");
            var nchecks = checks.length;
            var check;
            var data = "";
            for (var i = 0; i < nchecks; i++) {
                check = checks[i];
                if (check.checked) {
                    var id = check.parentNode.getAttribute("data-id");
                    if (id != null) {
                        data += id;
                        data += "-";
                        data += check.nextSibling.textContent;
                        data += "|";
                    }
                }
            }
            if (data.length > 0) data = data.substr(0, data.length - 1);
            return data;
        }

        this.obtenerChecksSeleccionados = obtenerChecksSeleccionados;
    }

    return GUI;
})();

var Validacion = (function () {
    function Validacion() {
    }
    Validacion.ValidarRequeridos = function (claseReq, div) {
        if (claseReq == null) claseReq = "R";
        /*  if (span == null) span = spnMensaje*/
        if (div == null) div = document;
        var controles = div.getElementsByClassName(claseReq);
        var nControles = controles.length;
        var c = 0;
        for (var i = 0; i < nControles; i++) {
            var tipocontrol = controles[i].tagName;
             
            switch (tipocontrol) {
                case "DIV":
                if (controles[i].innerHTML == "")
                {
                    controles[i].style.border = "red 1px solid";
                    c++
                } else
                {
                    controles[i].style.border = "";
                }
                    break;
                case "SPAN":
                    if (controles[i].innerHTML == "") {
                        controles[i].style.border = "red 1px solid";
                        c++
                    } else {
                        controles[i].style.border = "";
                    }
                    break;
                case "SELECT":
                    if (controles[i].value == "") {
                        controles[i].style.border = "red 1px solid";
                        c++
                    } else {
                        controles[i].style.border = "";
                    }
                    break;
                case "LABEL":
                    if (controles[i].innerHTML == "") {
                        controles[i].style.border = "red 1px solid";
                        c++
                    } else {
                        controles[i].style.border = "";
                    }
                    break;

                default:
                    if (controles[i].value == "") {
                        controles[i].style.borderColor = "red";
                        if (controles[i].id.substr(0, 3) == "cbs") {
                            controles[i].parentNode.style.border = "red 1px solid";
                        }
                        c++
                    } else {
                        controles[i].style.borderColor = "";
                        if (controles[i].id.substr(0, 3) == "cbs") {
                            controles[i].parentNode.style.border = "";
                        }
                    }
               
            }
            
            //if (controles[i].value == "") {
            //    controles[i].style.borderColor = "red";
            //    c++;
            //}
            //else {
            //    controles[i].style.borderColor = "";
            //}
        }
        /*  if (c > 0) span.innerHTML = "Los campos en Borde Rojo son Requeridos";*/
        if (c > 0) alerta([["Los campos en borde rojo son requeridos","advertencia"]]);
       /* else span.innerHTML = "";*/
        return(c == 0);
    }
    Validacion.ValidarNumeros = function (claseNum, span) {
        if (claseNum == null) claseNum = "N";
        if (span == null) span = spnMensaje
        var controles = document.getElementsByClassName(claseNum);
        var nControles = controles.length;
        var c = 0;
        for (var i = 0; i < nControles; i++) {
            if (isNaN(controles[i].value)) {
                controles[i].style.borderColor = "blue";
                c++;
            }
            else {
                controles[i].style.borderColor = "";
            }
        }
        if (c > 0) span.innerHTML = "Los campos en Borde Azul son Numeros";
        else span.innerHTML = "";
        return (c == 0);
    }
    Validacion.ValidarNumerosEnLinea = function (claseNum) {
        if (claseNum == null) claseNum = "N";
        var controles = document.getElementsByClassName(claseNum);
        var nControles = controles.length;
        for (var i = 0; i < nControles; i++) {
            controles[i].onkeyup = function (event) {
                var keycode = ('which' in event ? event.which : event.keycode);
                var esValido = ((keycode > 47 && keycode < 58) || (keycode > 95 && keycode < 106) || keycode == 8 || keycode == 37 || keycode == 39 || keycode == 110 || keycode == 188 || keycode == 190);
                if (!esValido) this.value = this.value.removeCharAt(this.selectionStart);
            }

            controles[i].onpaste = function (event) {
                event.preventDefault();
            }
        }
    }
    String.prototype.removeCharAt = function (i) {
        var tmp = this.split('');
        tmp.splice(i - 1, 1);
        return tmp.join('');
    }
    Validacion.ValidarDatos = function (claseReq, claseNum, span) {
        var valido = Validacion.ValidarRequeridos(claseReq);
        if (valido) {
            valido = Validacion.ValidarNumeros(claseNum, span);
        }
        return valido;
    }
    return Validacion;
})();

var FileSystem = (function () {
    function FileSystem() {
    }
    FileSystem.getMime = function (archivo) {
        var mime = "";
        var campos = archivo.split(".");
        var extension = campos[campos.length - 1].toLowerCase();
        switch (extension) {
            case "txt":
                mime = "text/plain";
                break;
            case "csv":
                mime = "text/csv";
                break;
            case "json":
                mime = "application/json";
                break;
            case "xlsx":
                mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                break;
            case "docx":
                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                break;
            case "pdf":
                mime = "application/pdf";
                break;
            default:
                mime = "application/octet-stream";
                break;
        }
        return mime;
    }
    FileSystem.download = function (data, archivo) {
        var mime = FileSystem.getMime(archivo);
        var blob = new Blob([data], { "type": mime });
        var link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = archivo;
        link.click();
    }
    return FileSystem;
})();

var Impresion = (function () {
    function Impresion() {
    }
    Impresion.imprimirTabla = function (tabla) {
        var ventana = window.frames["print_frame"];
        if (ventana != null) {
            var pagina = document.body;
            ventana.document.body.innerHTML = "";
            ventana.document.body.innerHTML = tabla;
            ventana.focus();
            ventana.print();
            ventana.close();
            document.body = pagina;
        }
    }
    Impresion.imprimirDiv = function (div) {
        var ventana = window.frames["print_frame"];
        if (ventana != null) {
            var pagina = document.body;
            mostrarControles(false);
            ventana.document.body.innerHTML = "";
            guardarValores(div);
            ventana.document.body.innerHTML = div.outerHTML;
            divVentana = ventana.document.getElementById(div.id);
            if (divVentana != null) recuperarValores(divVentana);
            ventana.focus();
            ventana.print();
            ventana.close();
            mostrarControles(true);
            document.body = pagina;
        }
    }

    function guardarValores(div) {
        if (div.hasChildNodes()) {
            var controles = div.childNodes;
            var ncontroles = controles.length;
            var control;
            for (var i = 0; i < ncontroles; i++) {
                control = controles[i];
                if (control.tagName == "INPUT" && control.type == "text") {
                    control.setAttribute("value", control.value);
                }
                guardarValores(control);
            }
        }
    }

    function recuperarValores(div) {
        if (div.hasChildNodes()) {
            var controles = div.childNodes;
            var ncontroles = controles.length;
            var control;
            for (var i = 0; i < ncontroles; i++) {
                control = controles[i];
                if (control.tagName == "INPUT" && control.type == "text") {
                    control.value = control.getAttribute("value");
                }
                recuperarValores(control);
            }
        }
    }

    function mostrarControles(visible) {
        var controles = document.getElementsByClassName("NoImprimir");
        var ncontroles = controles.length;
        var estilo = (visible ? "block" : "none");
        for (var j = 0; j < ncontroles; j++) {
            controles[j].style.display = estilo;
        }
    }
    return Impresion;
})();

var Popup = (function () {
    function Popup() {
    }
    Popup.CrearPopup = function (popup,div){
        var html = "";
        var nregistros = popup.length;
        var campos = [];
        var tipoControl = "";
        var esFilaOculta = "";
        if (ayudas == null) ayudas = [];
        for (var i = 1; i < nregistros; i++) {
            campos = popup[i].split("|");
            tipoControl = campos[1].substr(0, 3);
            esFilaOculta = (tipoControl=="hdf");
            html += "<div class='PopupWindowBoque grid_12'";
            if (esFilaOculta) html += " style='display:none'";
            html += ">";
            html += "<div class='texto'>";
            html += campos[0];
            html += "</div>";
            html += "<div >";
            switch (tipoControl) {
                case "hdf":
                    html += "<input type='hidden' id='";
                    html += campos[1];
                    html += "' class='";
                    html += campos[3];
                    html += "' data-campo='";
                    html += campos[4];
                    html += "'/>";
                    break;
                case "lbl":
                    html += "<label  id='";
                    html += campos[1];
                    html += "' class='";
                    html += campos[3];
                    html += "' data-campo='";
                    html += campos[4];
                    html += "'/>";
                    html += "</label>";
                    break;
                case "txt":
                    html += "<input type='text' id='";
                    html += campos[1];
                    html += "' style = 'max-width:";
                    html += campos[2];
                    html += "px' maxlength='";
                    html += campos[5];
                    html +="' class='";
                    html += campos[3];
                    html += "' data-campo='";
                    html += campos[4];
                    html += "'/>";
                    break;
                case "dtp":
                    html += "<input type='date' id='";
                    html += campos[1];
                    html += "' style = 'max-width:";
                    html += campos[2];
                    html += "px' ";
                    html += " class='";
                    html += campos[3];
                    html += "' data-campo='";
                    html += campos[4];
                    html += "'/>";
                    break;
                case "cbo":
                    html += "<select id='";
                    html += campos[1];
                    html += "' style = 'max-width:";
                    html += campos[2];
                    html += "px'";
                    html += "' class='";
                    html += campos[3];
                    html += "' data-campo='";
                    html += campos[4];
                    html += "'></select>"

                    break;
                //agregado por diego cbs combo seacrh
                case "cbs":
                    html += "<div style='max-width:";
                    html += campos[2];
                    html += "px' id='cbosearch";
                    html += campos[1];
                    html += "'>";
                    html += "<div class='content_seacrh_button'>";
                    html += "<div class='content_button'>";
                    html += "<span type='text' class='input_text";
                    html +="'></span> ";
                    html += "<div class='button_seacrh'>▲</div>";
                    html += "<input type='hidden' value='' id='";
                    html += campos[1];
                    html += "' class='";
                    html += campos[3];
                    html += "' data-campo='";
                    html += campos[4];
                    html += "'/>";
                    html += "</div>";
                    html += "<div class='content_seacrh'>";
                    html += " <input type='text' class='input_seacrh' placeholder='escriba aqui' />";
                    html += " <ul class='seacrh_list'>";
                    html += " </ul>";
                    html += "</div>";
                    html += "</div>";
                    html += "</div>";

            }
            html += "</div>";
            html += "</div>";
        }
        if (div != null) {
            div.innerHTML = html;
        } else {
            divPopup.innerHTML = html;
        }
     
    }

    Popup.Resize = function (popup, ancho, alto) {
        popup.style.width = ancho + "%";
        popup.style.height = alto + "%";
        popup.style.left = ((100 - ancho) / 2) + "%";
        popup.style.top = ((100 - alto) / 2) + "%";
    }

    Popup.ConfigurarArrastre = function(divPopupContainer, divPopupWindow, divBarra) {
        divBarra.draggable = true;
        divBarra.ondragstart = function (event) {
            var ancho = getComputedStyle(divPopupWindow, null).getPropertyValue("left");
            var alto = getComputedStyle(divPopupWindow, null).getPropertyValue("top");
            var a = Math.floor(ancho.replace("px", ""));
            var b = Math.floor(alto.replace("px", ""));
            var x = (event.clientX > a ? event.clientX - a : a - event.clientX);
            var y = (event.clientY > b ? event.clientY - b : b - event.clientY);
            var punto = x + "," + y;
            event.dataTransfer.setData("text", punto);
        }
        divBarra.ondragover = function (event) {
            event.preventDefault();
        }
        divPopupContainer.ondragover = function (event) {
            event.preventDefault();
        }
        divPopupContainer.ondrop = function (event) {
            event.preventDefault();
            var x1 = event.clientX;
            var y1 = event.clientY;
            var puntoInicial = event.dataTransfer.getData("text");
            var punto = puntoInicial.split(",");
            var x2 = punto[0] * 1;
            var y2 = punto[1] * 1;
            divPopupWindow.style.left = (x1 - x2) + "px";
            divPopupWindow.style.top = (y1 - y2) + "px";
        }
    }
    return Popup;
})();

var Speech = (function () {
    function Speech() {
    }
    Speech.Recognition = function() {
        if ('webkitSpeechRecognition' in window) {
            var recognizing = false;
            var recognition = new webkitSpeechRecognition();
            recognition.continuous = true;
            recognition.interimResults = false;
            iniciarReconocimiento();

            recognition.onresult = function (event) {
                var palabras = event.results[event.results.length - 1][0].transcript.trim();
                var palabra = palabras.split(" ");
                var comando = "";
                if (palabra.length > 0) {
                    comando = palabra[0];
                    var spnComando = document.getElementById("spnComando");
                    if (spnComando != null) spnComando.innerHTML = palabras;
                }
                ultimoComando = comando;
                ejecutarComandoVoz(palabras, comando);
            };

            recognition.onstart = function () {
                //alert("Iniciando Reconocimiento");
                recognizing = true;
            };

            recognition.onend = function () {
                //alert("Finalizando Reconocimiento");
                recognizing = false;
            };

            recognition.onerror = function (event) {
                alert(event.error);
            };

            function iniciarReconocimiento() {
                if (recognizing) {
                    recognition.stop();
                    return;
                }
                recognition.lang = "es-PE";
                recognition.start();
            }
        }
    }
    Speech.Synthesis = function(texto) {
        speechSynthesis.speak(new SpeechSynthesisUtterance(texto));
    }
    return Speech;
})();

var Chart = (function () {
    function Chart() {
    }

    Chart.dibujarTexto = function (contexto, fuente, color, texto, x, y, complejo) {
        if (complejo == null) contexto.beginPath();
        contexto.textBaseline = "top";
        contexto.font = fuente;
        contexto.fillStyle = color;
        contexto.fillText(texto, x, y);
        if (complejo == null) contexto.closePath();
    }

    Chart.dibujarRectangulo = function (contexto, x, y, ancho, alto, bordeAncho, bordeColor, colorRelleno) {
        contexto.beginPath();
        if (bordeAncho != null) contexto.lineWidth = bordeAncho;
        if (bordeColor != null) contexto.strokeStyle = bordeColor;
        contexto.rect(x, y, ancho, alto);
        if (colorRelleno != null) {
            contexto.fillStyle = colorRelleno;
            contexto.fill();
        }
        contexto.stroke();
        contexto.closePath();
    }

    Chart.dibujarLinea = function (contexto, x1, y1, x2, y2, ancho, color) {
        contexto.lineWidth = ancho;
        contexto.strokeStyle = color;
        contexto.moveTo(x1, y1);
        contexto.lineTo(x2, y2);
        contexto.stroke();
    }

    Chart.dibujarCirculo = function (contexto, x, y, radio, bordeAncho, bordeColor, colorRelleno, complejo) {
        if (complejo == null) contexto.beginPath();
        if (bordeAncho != null) contexto.lineWidth = bordeAncho;
        if (bordeColor != null) contexto.strokeStyle = bordeColor;
        contexto.arc(x, y, radio, 0, 2 * Math.PI, false);
        if (colorRelleno != null) {
            contexto.fillStyle = colorRelleno;
            contexto.fill();
        }
        contexto.stroke();
        if (complejo == null) contexto.closePath();
    }

    Chart.dibujarTextoVertical = function (contexto, fuente, color, texto, x, y) {
        contexto.save();
        contexto.translate(x, y);
        contexto.rotate(-Math.PI / 2);
        contexto.font = fuente;
        contexto.fillStyle = color;
        contexto.fillText(texto, 0, 0);
        contexto.restore();
    }

    Chart.dibujarBarra3D = function (contexto, x, y, ancho, alto, bordeAncho, bordeColor, colorRelleno, profundidad) {
        contexto.beginPath();
        if (bordeAncho != null) contexto.lineWidth = bordeAncho;
        if (bordeColor != null) contexto.strokeStyle = bordeColor;
        contexto.rect(x, y, ancho, alto);
        contexto.lineWidth = bordeAncho;
        contexto.strokeStyle = bordeColor;
        contexto.moveTo(x, y);
        contexto.lineTo(x + profundidad, y - profundidad);
        contexto.lineTo(x + ancho + profundidad, y - profundidad);
        contexto.lineTo(x + ancho + profundidad, y + alto - profundidad);
        contexto.lineTo(x + ancho, y + alto);
        contexto.moveTo(x + ancho + profundidad, y - profundidad);
        contexto.lineTo(x + ancho, y);
        contexto.stroke();
        if (colorRelleno != null) {
            contexto.fillStyle = colorRelleno;
            contexto.fill();
        }
        contexto.stroke();
        contexto.closePath();
    }

    Chart.dibujarCurva = function (contexto, x1, y1, x2, y2, ancho, color) {
        contexto.lineWidth = ancho;
        contexto.strokeStyle = color;
        contexto.moveTo(x1, y1);
        contexto.bezierCurveTo(x1, y2, x2, y2, x2, y2);
        contexto.stroke();
    }

    Chart.dibujarLineaSuperficie = function (contexto, x1, y1, x2, y2, ancho, color) {
        contexto.lineWidth = ancho;
        contexto.strokeStyle = color;
        contexto.lineTo(x2, y2);
    }

    return Chart;
})();
window.Chart = Chart;



var formulario = (function () {
    function formulario() {
    }
    formulario.CrearFormulario = function (formulario,div) {
        var html = "";
        var nregistros = formulario.length;
        var campos = [];
        var tipoControl = "";
        var esFilaOculta = "";
        var bloques = formulario[1];
        var camposBloques = bloques.split('|');
        var nbloques = camposBloques.length;
        html += "<div class='form_basic_horizontal grid_12 grid_start p0'>";
        for (var i = 0; i < nbloques; i++) {
            html += "<div class='form_basic_horizontal grid_12 grid_start p0'>";
            for (var j= 2; j < nregistros; j++) {
                campos = formulario[j].split("|");
                if (campos[6] == camposBloques[i]) {
                    tipoControl = campos[1].substr(0, 3);
                    esFilaOculta = (tipoControl == "hdf");

                    if (tipoControl == "swt" || tipoControl== "btn") {
                        html += "<div class='form_basic_horizontal_permanent ";
                    }
                    else {
                        html += "<div class='";
                        switch (campos[7]) {
                            case "grid_6":
                                html += "form_2_column_6 "
                                break;
                            case "grid_12":
                                html += "form_2_column "
                                break;
                            case "grid_8":
                                html += "form_2_column_8 "
                                break;
                            case "grid_4":
                                html += "form_2_column_4 "
                                break;
                            case "grid_3":
                                html += "form_2_column_4 "
                                break;
                            default:
                                html += "form_2_column ";
                        }
                     
                    }
                   
                    html += campos[7];
                    html +=" ";
                    html += campos[8];
                        html += "'";
                        if (esFilaOculta) html += " style='display:none'";
                    html += ">";
                    if (tipoControl != "swt" && tipoControl != "btn") {
                        html += "<div class='texto'>";
                        html += campos[0];
                        html += "</div>";                      
                    }

                    switch (tipoControl) {
                        case "hdf":
                            html += "<input type='hidden' id='";
                            html += campos[1];
                            html += "' class='";
                            html += campos[3];
                            html += "' data-campo='";
                            html += campos[4];
                            html += "'/>";
                            break;
                        case "txt":
                            html += "<input type='text' id='";
                            html += campos[1];
                            html += "' style='max-width:";
                            html += campos[2];
                            html += "px' maxlength='";
                            html += campos[5];
                            html += "' class='";
                            html += campos[3];
                            html += "' data-campo='";
                            html += campos[4];
                            html += "'/>";
                            break;
                        case "btn":
                            html += "<input type='button' id='";
                            html += campos[1];
                            html += "'";
                            html += "value='";
                            html += campos[0];
                            html += "'  style='width:";
                            html += campos[2];
                            html += "px' data-campo='";
                            html += campos[4];
                            html += "' class='";
                            html += campos[3];
                            html += "'/>";
                            break;
                        case "cbo":
                            html += "<select id='";
                            html += campos[1];
                            html += "' style = 'margin: 5px;max-width:";
                            html += campos[2];
                            html += "px'";
                            html += "' class='";
                            html += campos[3];
                            html += "' data-campo='";
                            html += campos[4];
                            html += "'></select>"

                            break;
                        //agregado por diego cbs combo seacrh
                        case "cbs":
                            html += "<div id='cbosearch";
                            html += campos[1];
                            html += "' class='grid_12'>";
                            html += "<div class='content_seacrh_button'>";
                            html += "<div class='content_button'>";
                            html += "<span type='text' class='input_text";
                            html += "'></span> ";
                            html += "<div class='button_seacrh'>▲</div>";
                            html += "<input type='hidden' value='' id='";
                            html += campos[1];
                            html += "' class='";
                            html += campos[3];
                            html += "' data-campo='";
                            html += campos[4];
                            html += "'/>";
                            html += "</div>";
                            html += "<div class='content_seacrh'>";
                            html += " <input type='text' class='input_seacrh' placeholder='escriba aqui' />";
                            html += " <ul class='seacrh_list'>";
                            html += " </ul>";
                            html += "</div>";
                            html += "</div>";
                            html += "</div>";
                            break;
                        case "tme":
                            html += "<input type='time' id='";
                            html += campos[1];
                            html += "' style = 'max-width:";
                            html += campos[2];
                            html += "px' maxlength='";
                            html += campos[5];
                            html += "' class='";
                            html += campos[3];
                            html += "' data-campo='";
                            html += campos[4];
                            html += "'/>";

                            break;
                        case "txa":
                            html += "<textarea id='";
                            html += campos[1];
                            html += "' style = 'max-width:";
                            html += campos[2];
                            html += "px' rows='6'  maxlength='";
                            html += campos[5];
                            html += "' class='";
                            html += campos[3];
                            html += "' data-campo='";
                            html += campos[4];
                            html += "'></textarea>"
                            break;
                        case "swt":
                            html += " <div class='form_basic_horizontal  grid_center grid_12'>";
                            html += " <div class='navigation_nav_last'>";
                            html += campos[0];
                            html += "</div >";
                            html += "<div class='switch'>";
                            html += "   <input id='";
                            html += campos[1];
                            html += "' class='switch-input ";
                            html += campos[3];
                            html += "' type='checkbox' />";
                            html += "   <label for='";
                            html += campos[1];
                            html += "' class='switch-label'></label>";
                            html += " </div>";
                            html += " <div class='navigation_nav_last'>";
                            html += campos[4];
                            html += "</div>";
                            html += "</div>";
                            break;

                        case "div":
                            html += "<div id='";
                            html += campos[1];
                            html += "' style = 'max-width:";
                        /*    html += campos[2];*/
                            html += "px'";
                            html += "' class='texto ";
                            html += campos[3];
                            html += "' data-campo='";
                            html += campos[4];
                            html += "'></div>"

                    }
                 
                    html += "</div>";
                }
            }
            html += "</div>";
        }
        html += "</div>";
        if (ayudas == null) ayudas = [];
        div.innerHTML = html;
     
    }
    return formulario;
})();



function crearNavigationNav(idNavigationNav,id) {
    var html = "";
    html += "<img src='";
    html += hdfRaiz.value;
    html+= "Images/home.svg' class='icono_navigation_nav' />";
    html += "<div class='navigation_nav'>Home</div>";
    var navigation = [];
    if (sessionStorage.getItem('navigation_page' + id) != null) {
        navigation = sessionStorage.getItem('navigation_page' + id).split('|');
    }else {
        navigation = sessionStorage.getItem('navigation_page').split('|');
        sessionStorage.setItem('navigation_page' + id, sessionStorage.getItem('navigation_page'));
    };
    var navcampos = navigation ;
    var ncampos = navcampos.length;
    for (var i = 0; i < ncampos; i++) {
        if ((ncampos - 1) == i) {
            html += "<div class='navigation_nav_last'>";
            html += "/ " + navcampos[i];
            html += "</div>";
        }
        else {
            html += "<div class='navigation_nav'>";
            html += "/ " + navcampos[i];
            html += "</div>";
        }

    }
   
    idNavigationNav.innerHTML = html;

}


async function alerta(mensaje, cerrarAccion = false) {
    tipo = {
        exito: { tipoAlerta: 'exito', color: '#28a745', class: 'exito', img: hdfRaiz.value + 'Images/success.svg', },
        advertencia: { tipoAlerta: 'advertencia', color: '#ffc107', class: 'advertencia', img: hdfRaiz.value + 'Images/warning.svg', },
        error: { tipoAlerta: 'error', color: '#dc3545', class: 'error', img: hdfRaiz.value + 'Images/error.svg', }
    };
    var html = "";
    html += "<div  id='alertPopup' class='PopupContainerAlert' style='display: flex; justify-content: flex-end; align-items:flex-end;flex-wrap: wrap; overflow: scroll;overflow-x: hidden;'>";
    if (typeof(mensaje) == "object" && typeof(mensaje[0]) == "string") {
        html += crear(tipo[mensaje[1]].class, tipo[mensaje[1]].img, mensaje[0]);    
    }
    html += "</div>";
    document.body.insertAdjacentHTML('beforeend', html);
    var nmensaje = mensaje.length-1;

    //if (typeof(mensaje) == "object" && typeof(mensaje[0]) =="object") {
    //    var interval = 400;
    //    mensaje.forEach((mensaje, i) => {
    //        setTimeout(() => {
    //            document.getElementById("alertPopup").insertAdjacentHTML('beforeend', crear(tipo[mensaje[1]].class, tipo[mensaje[1]].img, mensaje[0]));
                
    //        }, i * interval)
    //    })
    //}
    this.cerraralert = function (input) {
        var popupAlert = input.parentNode.parentNode.parentNode;
        popupAlert.children[0].classList.add('desaparecer');
        setTimeout(() => {
            document.body.removeChild(popupAlert);
            if (cerrarAccion) cerraralertaPopud();
        }, 200)
    }
    await crearAlertars();

    async function crearAlertars() {
        return new Promise(resolve => {
            if (typeof (mensaje) == "object" && typeof (mensaje[0]) == "object") {
                var interval = 400;
                mensaje.forEach((mensaje, i) => {
                    setTimeout(() => {
                        document.getElementById("alertPopup").insertAdjacentHTML('beforeend', crear(tipo[mensaje[1]].class, tipo[mensaje[1]].img, mensaje[0]));
                        if ((nmensaje) * 400 == i * interval) {
                            resolve(true);
                        }
                    }, i * interval)
                })
            }
        });
    }


   var alertas = document.getElementById("alertPopup").children;
   var alertaPopud = document.getElementById("alertPopup");
  await  delayAlerta(alertas);

    async function delayAlerta(listaAlertas) {
        var nListasAlertas = listaAlertas.length;
        for (var i = 0; i < nListasAlertas; i++) {
            if (listaAlertas[0] == null) {
                break;
            }
            else {
                await removeAlertars(listaAlertas[i]);
            }
            
            i--;
            
        }
        document.body.removeChild(document.getElementById("alertPopup"));    
    }

    async function removeAlertars(item) {
        return new Promise(resolve => {
           item.classList.add('desaparecer');
            setTimeout(() => {
                alertaPopud.removeChild(item);
                resolve(true);
            }, 1500)
        });
    }
 
    function crear(clase, img, mensaje) {
        var pop = "";
        pop += "<div class='PopupWindowAlert1 ";
        pop += clase;
        pop += "'>";
        pop += "<div   class='title'>";
        pop += "<img style='width:30px;height:auto;'  src='";
        pop += img;
        pop += "'>";
        pop += "</div>";
        pop += "<div  class='titleAlert'>";
        pop += mensaje;
        pop += "</div>";
        pop += "<div class='divPopupAccion' >";
        pop += "<input type='button' value='X'  class='x p0' onclick='cerraralert(this);' style='color:";
        pop += "'/> ";
        pop += "</div>";
        pop += "</div>";
        return pop;
    }

}

function mostrarloader() {
    var html = "";
    html += "<div id='loader' class='loader_back'>";
    html += "<div class='loader_center'>";
    html += "<div class='loader_ring'></div>";
    html += "<span></span>";
    html += "</div>";
    html += "</div>";
    document.body.insertAdjacentHTML('beforeend', html);
}

function cerrarloader() {
   var loader= document.getElementById("loader");
    document.body.removeChild(loader);

}

function crearConfirm(mensaje,MensajeAlerta) {
    var confirmacion ;
    var html = "";
    html += "<div id='idConfirmPopup' class='PopupContainer' style='display:inline;'>";
    html += "<div class='PopupWindow' style = '    border-radius: 15px;width: 300px; height: auto; margin: auto; margin-top: 5%; position: initial; color: white; background-color:var(--color-muniz-primary-verde) ; font-weight: bold;    overflow: hidden;'>";
    html += "<div class='Cuerpo' style='text-align: left;overflow: hidden;'>";
    html += mensaje;
    html += "</div>";
    html += "<div style='border: 0px; background-color: whitesmoke; display: flex; flex-wrap: nowrap; align-items: flex-end; justify-content: flex-end; text-align: center;'>";
    html += "<input value='SI' type='button' class='grid_3' />";
    html += "<input value='NO' type='button' class='grid_3 gris' />";

        html += "</div>";
        html += "</div>";
        html += "</div>";
        document.body.insertAdjacentHTML('beforeend', html);
 
    var confirm = document.getElementById("idConfirmPopup");
    var btnAceptar = confirm.children[0].children[1].children[0];
    var btnCancelar = confirm.children[0].children[1].children[1];
    btnCancelar.onclick = function () {
        console.log("btnCancelar");
        confirmacion = false;
    }
    btnAceptar.onclick = function () {
        console.log("btnAceptar");
        confirmacion = true;
    }

    this.obtenerConfirmacion=  function() {
        return new Promise(resolve => {
            var intervalo = setInterval(function () {
                if (confirmacion == false || confirmacion == true) {
                    clearInterval(intervalo);
                    document.body.removeChild(confirm);
                    resolve(confirmacion);
                }
            }, 10);
        });
    }
 
}

function compareDosFechas(fechaIni, fechaFin) {
    var exito = false;
    if (fechaIni.getTime() < fechaFin.getTime()) exito = true;
    return exito;
}

function validaResponseData(rpta) {
    var valida = true;
    if (rpta.substring(0, 5) == "ERROR") {
        alerta(["Hubo un error intenta nuevamente", 'advertencia']);
        valida = false;
    }
    if (rpta == "") {
        alerta(["Hubo un error intenta nuevamente", 'error']);
        valida = false;
    }
    if (rpta == null) {
        alerta(["Hubo un error intenta nuevamente", 'error']);
        valida = false;
    }
    return valida;
}

function html_escape(text) {
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
function html_escape_envioDB(text) {
    return text
        .replace(/'/g, "''")
}

function crearResponsiveTable(listaCabecera, tabla, cabeceraNo) {
    var hoja = document.createElement("style");
    var style = "";
    style += "@media (max-width: 875px){";
    style += " #" + tabla + " > table              { display:block;}";
    style += " #" + tabla + "  *  { display:block ;}";
    style += " #" + tabla + "  tr  { display:block ;}";
    style += " #" + tabla + "  td  { display:block ;}";
    style += " #" + tabla + "  th   { display:block ;}";
    style += " #" + tabla + " >table>thead>tr       {display:flex;justify-content: space-around; flex-wrap: wrap;align-items: center;}";
    if (cabeceraNo) {
        style += " #" + tabla + ">table>thead>tr>th          { display:inline-block;padding:2px;}";
    }
    else {
        style += " #" + tabla + ">table>thead>tr>th          { display:none;}";
    }

    style += " .ocultar             {display:none !important;}";
    style += " #" + tabla + ">table>tbody>tr             { height:auto;}";
    style += " #" + tabla + ">table>tbody>tr>td                   {display: flex;flex-wrap: nowrap;align-items: center; justify-content: flex-start; text-align: left; }";
    style += " #" + tabla + ">table>tbody>tr>td:last-child        {margin-bottom:0 }";
    style += " #" + tabla + ">table>tfoot>tr            {display: flex;justify-content:space-around;}";
    style += " #" + tabla + ">table>tfoot>tr>td           { display: flex;flex-wrap: wrap;width: 100 %;justify-content: center;}";
    style += " #" + tabla + ">table>tbody>tr>td:before            {left:10px;font-weight:700;width:40%;color:var(--color-muniz-primary-verde);}";
    style += ".Mostrar {display:inline !important;}";
    for (var i = 0; i < listaCabecera.length; i++) {
        style += " #tbData";
        style += tabla;
        style += " >tr>td:nth-child(";
        style += i + 1;
        style += "):before { content:'";
        style += listaCabecera[i];
        style += "';}";
    }
    style += " } ";
    hoja.innerHTML = style;
    document.body.appendChild(hoja);
}



