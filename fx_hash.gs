/**
 * Calcula el hash de las celdas pasadas como parámetros.
 * @return matriz de hashes calculados
 * @param {A1:D10} valores Datos de entrada como texto.
 * @param {verdadero} base64 Codificar en Base64 (VERDADERO | FALSO).
 * @param {"SHA256"} tipo Tipo de hash (MD2| MD5 | SHA1 | SHA256 | SHA384 | SHA512).
 *
 * @return matriz de hashes calculados codificados en base64, si procede.
 *
 * @customfunction
 *
 * (c) Pablo Felip @pfelipm GNU GPL v3
 */ 
 
function HASH(valores, base64, tipo) {

  var algoritmo;
  
  // Comprobación de parámetros
  
  if (valores == '') {return "Rango de datos no especificado";}
  if (typeof base64 != 'boolean') {return "Base64 sí/no no especificado";}
  if (!tipo) {return "Tipo de hash no especificado";}

  // Tipo de hash

  switch (tipo.toUpperCase()) {
    case 'MD2': algoritmo = Utilities.DigestAlgorithm.MD2; break;
    case 'MD5': algoritmo = Utilities.DigestAlgorithm.MD5; break;
    case 'SHA1': algoritmo = Utilities.DigestAlgorithm.SHA_1; break;
    case 'SHA256': algoritmo = Utilities.DigestAlgorithm.SHA_256; break;
    case 'SHA384': algoritmo = Utilities.DigestAlgorithm.SHA_384; break;
    case 'SHA512': algoritmo = Utilities.DigestAlgorithm.SHA_512; break;
    default: return 'Tipo de hash desconocido';
  }
  
  // Generar hashes
  
  if (valores.map) {
  
    // Vector o matriz
    
    return valores.map(function(c) {  
      return c.map(function(c) {
      
        // Se recodifica en Base64
      
        if (base64) {return Utilities.base64Encode(Utilities.computeDigest(algoritmo,c, Utilities.Charset.UTF_8));}
        
        // ... o se recodifica como cadena hexadecimal
        
        else {return bin2hex(Utilities.computeDigest(algoritmo,c, Utilities.Charset.UTF_8));}
      });
    });
  }
  else {
  
    // Valor único, se recodifica en Base64
    
    if (base64) {return Utilities.base64Encode(Utilities.computeDigest(algoritmo,valores, Utilities.Charset.UTF_8));}
    
    // ... o se recodifica como cadena hexadecimal
    
    else {return bin2hex(Utilities.computeDigest(algoritmo,valores, Utilities.Charset.UTF_8));}
  } 

}

/**
 * Recodifica en base64 el contenido de las celdas pasadas como parámetros.
 * @return matriz de texto condificado en base64
 * @param {A1:D10} valores Datos de entrada como texto.
 *
 * @return matriz de hashes calculados codificados en base64, si procede.
 *
 * @customfunction
 *
 * (c) Pablo Felip @pfelipm GNU GPL v3
 */ 

function BASE64(valores) {

  if (!valores) {return "Rango de datos no especificado";}
  else {
  
    if (valores.map) {
    
      // Vector o matriz

      valores = valores.map(function(c) {
      
         return c.map(function(c) {return Utilities.base64Encode(c, Utilities.Charset.UTF_8);    
        })
      })
      return valores;
    }
    else {return Utilities.base64Encode(valores, Utilities.Charset.UTF_8);}    
  }
  
}