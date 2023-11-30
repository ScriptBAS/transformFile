function importExcelData() {
	var file = document.getElementById("excelFile").files[0];
	var reader = new FileReader();
	reader.onload = function (e) {
		var data = new Uint8Array(e.target.result);
		var workbook = XLSX.read(data, { type: "array" });
		var worksheet = workbook.Sheets[workbook.SheetNames[0]];
		var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
		// Теперь вы можете обработать данные Excel (например, создать таблицу)
		let newData = transformData(jsonData);
		var ws = XLSX.utils.aoa_to_sheet(newData);
		XLSX.utils.book_append_sheet(workbook, ws, "Сварные швы");
		XLSX.writeFile(workbook, file.name, { compression: true });
	};
	reader.readAsArrayBuffer(file);
}

function transformData(jsonData) {
	const newTable = {};
	const thickness = {};
	const errors = [];
	for (i = 2; i < jsonData.length; i++) {
		let current = jsonData[i];
		if (!current[2]) break;
		let elements = current[1].trim().split("-");
		if (elements[0] === "0") continue;
		if (!newTable[elements[0]]) {
			newTable[elements[0]] = current[4];
			thickness[elements[0]] = Math.ceil(current[8]);
		} else {
			thickness[elements[0]] = Math.min(
				thickness[elements[0]],
				Math.ceil(current[8])
			);
			if (newTable[elements[0]] !== current[4]) {
				errors.push(
					`Несоответствие диаметра для сварного шва ${elements[0]}`
				);
			}
		}
		if (elements[1] === "0") continue;
		if (!newTable[elements[1]]) {
			thickness[elements[1]] = Math.ceil(current[8]);
			let diametr = current[2] === "п" ? current[5] : current[4];
			newTable[elements[1]] = diametr;
		} else {
			thickness[elements[1]] = Math.min(
				thickness[elements[1]],
				Math.ceil(current[8])
			);
			if (newTable[elements[1]] !== diametr) {
				errors.push(
					`Несоответствие диаметра для сварного шва ${elements[1]}`
				);
			}
		}
	}
	printErrors(errors);
	const data = [];
	for (key in newTable) {
		data[data.length] = [key, newTable[key], thickness[key]];
	}
	return data.sort((a, b) => {
		return a[0].split("/")[0] - b[0].split("/")[0];
	});
}

function printErrors(errors) {
	const errorsTable = document.getElementById("errors");
	errors.forEach((err) => {
		const newError = document.createElement("div");
		newError.innerHTML = err;
		errorsTable.append(newError);
	});
}
