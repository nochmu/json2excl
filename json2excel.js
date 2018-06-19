#!/usr/bin/env node

const program = require('commander');
const XlsxPopulate = require('xlsx-populate');
const log = require('./src/logger');
const util = require('util');


function parse_json_file(file){
	var json = JSON.parse(require('fs').readFileSync(file, 'utf8'));
	return json;
}

function parse_template_file(file){
	log.push("parse template");
	log.trace("file: %s", file);
	var tmpl = parse_json_file(file);
	log.trace("content:");
	log.indent("%s", util.inspect(tmpl,  { showHidden: false, depth: null , compact: false}));
	log.pop();
	return tmpl;
}

function parse_data_file(file){
	log.push("parse json");
	log.trace("file: %s", file);
	var data = parse_json_file(file);
	log.trace("content:");
	log.indent("%s", util.inspect(data,  { showHidden: false, depth: null , compact: false}));
	log.pop();
	return data;
}

function apply_template(sheet, tmpl){
	log.push("apply template");

	if (typeof tmpl.columns !== "undefined") {
		for(col of tmpl.columns){
			log.push("column:%s", col.id);
			log.trace("style:%o", col.style);
			sheet.column(col.id).style(col.style);
			log.pop();
		}
	}

	log.pop();
}

function generate_excel(tmpl, dat){
	var excl = XlsxPopulate.fromBlankAsync().then(workbook => {
		log.push("generate workbook");
		// Modify the workbook.
		var sheet = workbook.sheet("Sheet1");
		log.trace("sheet: %o", sheet.name());
		var r = sheet.cell("A1").value(dat);
		apply_template(sheet, tmpl);

		log.pop();
		return workbook;
	});

	return excl;
}



function write_excel(excl, file){
	var r = excl.then(wb => {
		log.push("write workbook");
		log.trace("file: %s", file);
		var x = wb.toFileAsync(file);
		log.pop();
		return x;
	});
	return r;
}

// ------------------------------------ MAIN
program
  .usage('[options]')
  .version('0.0.1', '-v, --version')
  .option('-j, --json <file>', 'JSON file input')
  .option('-t, --template <file>', 'Template file')
  .option('-o, --output <file>', 'Output file')
  .option('-V, --verbose')
  .parse(process.argv);

if (program.verbose) log.setLevel("trace");

var template = parse_template_file(program.template);
var data = parse_data_file(program.json);

var excel = generate_excel(template, data);
write_excel(excel, program.output);

