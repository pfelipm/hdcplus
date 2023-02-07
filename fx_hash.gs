/**
 * Calcula el hash de los valores del intervalo.
 * @param {A1:D10} valores Datos de entrada como texto.
 * @param {verdadero} base64VoF Codificar en Base64 (VERDADERO | FALSO).
 * @param {"SHA256"} tipo Tipo de hash (MD2 | MD5 | SHA1 | SHA256 | SHA384 | SHA512).
 *
 * @return matriz de hashes calculados codificados en base64, si procede.
 *
 * @customfunction
 *
 * (c) Pablo Felip @pfelipm GNU GPL v3
 */ 
 
function HASH(valores, base64VoF, tipo) {

  var algoritmo;
  
  // Comprobación de parámetros

  if (typeof valores == 'undefined' || typeof base64VoF == 'undefined' || typeof tipo == 'undefined')
    {return '!Faltan argumentos';}
  if (typeof base64VoF != 'boolean') {return "!Base64 V/F inválido";}
  if (typeof tipo != 'string') {return "!Tipo de hash inválido";}

  // Los parámetros parecen correctos ¡adelante!

  switch (tipo.toUpperCase()) {
    case 'MD2': algoritmo = Utilities.DigestAlgorithm.MD2; break;
    case 'MD5': algoritmo = Utilities.DigestAlgorithm.MD5; break;
    case 'SHA1': algoritmo = Utilities.DigestAlgorithm.SHA_1; break;
    case 'SHA256': algoritmo = Utilities.DigestAlgorithm.SHA_256; break;
    case 'SHA384': algoritmo = Utilities.DigestAlgorithm.SHA_384; break;
    case 'SHA512': algoritmo = Utilities.DigestAlgorithm.SHA_512; break;
    default: return '!Tipo de hash inválido';
  }
  
  // Generar hashes
  
  if (valores.map) {
  
    // Vector o matriz
    
    return valores.map(function(c) {  
      return c.map(function(c) {
      
        // Se recodifica en Base64
      
        if (base64VoF) {return Utilities.base64Encode(Utilities.computeDigest(algoritmo,c, Utilities.Charset.UTF_8));}
        
        // ... o se recodifica como cadena hexadecimal
        
        else {return bin2hex(Utilities.computeDigest(algoritmo,c, Utilities.Charset.UTF_8));}
      });
    });
  }
  else {
  
    // Valor único, se recodifica en Base64
    
    if (base64VoF) {return Utilities.base64Encode(Utilities.computeDigest(algoritmo,valores, Utilities.Charset.UTF_8));}
    
    // ... o se recodifica como cadena hexadecimal
    
    else {return bin2hex(Utilities.computeDigest(algoritmo,valores, Utilities.Charset.UTF_8));}
  } 

}

/**
 * Recodifica en base64 el contenido del intervalo.
 * @param {A1:D10} valores Datos de entrada como texto.
 * @param {falso} esHexa El valor a recodificar es una cadena binaria
                         representada como texto hexadecimal (VERDADERO | FALSO).
                         Si no se especifica toma el valor FALSO.
 * @return matriz de valores codificados en base64.
 *
 * @customfunction
 *
 * (c) Pablo Felip @pfelipm GNU GPL v3
 */ 

function BASE64(valores, esHexa) {

  // Comprobación de parámetros

  if (typeof valores == 'undefined') {return '!Faltan argumentos';}
  if (typeof esHexa == 'undefined') {esHexa = false;}  
  if (typeof esHexa != 'boolean') {return 'esHexa inválido';}  

  // Los parámetros parecen correctos ¡adelante!
  
  if (valores.map) {
    
    // Vector o matriz
    
    valores = valores.map(function(c) {
      
      if (!esHexa) {
        
        return c.map(function(c) {return Utilities.base64Encode(c, Utilities.Charset.UTF_8);})
      }
      else {
        
        return c.map(function(c) {return Utilities.base64Encode(Utilities.newBlob(hex2bytes(c).map(function(byte){
        
          return byte <128 ? byte : byte - 256;})).getBytes());})
      }
    })
    return valores;
  }
  else {
    if (!esHexa) {
      
      return Utilities.base64Encode(valores, Utilities.Charset.UTF_8);
    }    
    else {
        
      return Utilities.base64Encode(Utilities.newBlob(hex2bytes(valores).map(function(byte){
        
        return byte <128 ? byte : byte - 256;})).getBytes());
    }
  }      
  
}