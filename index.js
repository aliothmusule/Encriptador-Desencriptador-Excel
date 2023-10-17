const XLSX = require('xlsx');
const bcrypt = require('bcryptjs');
const fs = require('fs');

async function encryptPasswords() {
  try {
    const originalWorkbook = XLSX.readFile('./BASE DE DATOS-CORREOS NUEVO INGRESO_Correomascontra_jace.xlsx');
    const worksheetName = 'Hoja2';
    const columnToEncrypt = 'L';
    const newColumn = 'M';

    const originalWorksheet = originalWorkbook.Sheets[worksheetName];
    const range = XLSX.utils.decode_range(originalWorksheet['!ref']);

    const newWorkbook = XLSX.utils.book_new(); // Crear un nuevo libro de trabajo
    const newWorksheet = XLSX.utils.json_to_sheet([]); // Crear una nueva hoja de cálculo vacía

    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'ContraseñasEncriptadas'); // Agregar la hoja de cálculo al nuevo libro de trabajo

    const totalRows = range.e.r - range.s.r;

    // Definir una animación de espera
    const animationFrames = ['|', '/', '-', '\\'];
    let currentFrameIndex = 0;

    // Iniciar la animación de espera
    const animationInterval = setInterval(() => {
      process.stdout.write(`\rProcesando: ${animationFrames[currentFrameIndex]}`);
      currentFrameIndex = (currentFrameIndex + 1) % animationFrames.length;
    }, 100);

    for (let i = range.s.r + 1; i <= range.e.r; i++) {
      const cellAddress = `${columnToEncrypt}${i + 1}`;
      const cell = originalWorksheet[cellAddress];

      if (cell && cell.v) {
        const plaintextPassword = cell.v.toString();
        const hashedPassword = await bcrypt.hash(plaintextPassword, 10);

        // Agregar la contraseña encriptada al nuevo libro de trabajo
        XLSX.utils.sheet_add_json(newWorksheet, [{ [newColumn]: hashedPassword }], { skipHeader: true, origin: -1 });

        // Imprimir la contraseña encriptada en la consola
        console.log(`Contraseña encriptada en la fila ${i + 1}: ${hashedPassword}`);
      }

      // Calcular el progreso y actualizar la consola
      const progress = ((i - range.s.r) / totalRows) * 100;
      process.stdout.write(`\rProcesando: ${animationFrames[currentFrameIndex]} ${progress.toFixed(2)}%`);
    }

    clearInterval(animationInterval); // Detener la animación de espera

    // Guardar el nuevo libro de trabajo con las contraseñas encriptadas
    XLSX.writeFile(newWorkbook, 'archivo_encriptado.xlsx');
    console.log('\nContraseñas encriptadas y guardadas en el archivo "archivo_encriptado.xlsx"');
  } catch (error) {
    clearInterval(animationInterval); // Detener la animación en caso de error
    console.error('Error al procesar el archivo de Excel:', error.message);
  }
}

encryptPasswords();
