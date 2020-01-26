/**
 * Divide en subcadenas los valores de *texto* contenidos en las celdas
 * de un intervalo utilizando una secuencia de caracteres delimitadora (separador),
 * que se omite en las subcadenas resultantes, y las organiza en un vector columna.
 *
 * @param {B2:B20} intervalo
 * Intervalo con valores de texto cuyo contenido se desea trocear.
 * 
 * @param {", "} separador
 * Subcadena empleada como delimitador.
 *
 * @return vector columna
 *
 * @customfunction
 *
 * (c) Pablo Felip @pfelipm GNU GPL v3
 */

function TROCEAR(intervalo, separador) {
  
  var i, p;
  var resultado = [];
  
  // Comprobación de parámetros
 
  if (typeof intervalo == 'undefined' || typeof separador == 'undefined') {return '!Faltan argumentos';}
  if (!intervalo) return '!No hay valores a trocear';
  if (typeof separador != 'string' || !separador) return '!Separador inválido';

  // Los parámetros parecen correctos ¡adelante!

  if (intervalo.map) {
  
    // Parámetro es vector o matriz
    
    intervalo.map(function(d) {
            
      d.map(function(d) {
      
        if (typeof d == 'string') {
          i = 0;
          p = d.indexOf(separador, 0);
          while (p != -1) {        
            if (p > 0) {resultado.push (d.substring(i, p));} // por si cadena comienza por separador
            i = p + separador.length;
            p = d.indexOf(separador, p + separador.length);           
          }
          if (i < d.length)  {resultado.push (d.substring(i, d.length));} // último fragmento
        }
      })
    })
   if (resultado.length > 0){ return resultado; }
   else {return null;}
  }
  
  else {
  
    // Parámetro es valor único
    
    if (typeof intervalo == 'string') {
      i = 0;
      p = intervalo.indexOf(separador, 0);
      while (p != -1) {        
        if (p > 0) {resultado.push (intervalo.substring(i, p));}  // por si cadena comienza por separador
        i = p + separador.length;
        p = intervalo.indexOf(separador, p + separador.length);           
      }
      if (i < intervalo.length) {resultado.push (intervalo.substring(i, intervalo.length));} // último fragmento 
      
      return resultado;
    }
  }

}