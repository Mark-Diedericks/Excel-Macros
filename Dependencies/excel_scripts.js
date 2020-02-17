Blockly.Python['excel_cells'] = function(block) {
  var value_row = Blockly.Python.valueToCode(block, 'Row', Blockly.Python.ORDER_ATOMIC);
  var value_column = Blockly.Python.valueToCode(block, 'Column', Blockly.Python.ORDER_ATOMIC);
  // TODO: Assemble Python into code variable.
  var code = '...';
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.Python.ORDER_NONE];
};

Blockly.Python['excel_range'] = function(block) {
  var value_cell1 = Blockly.Python.valueToCode(block, 'Cell1', Blockly.Python.ORDER_ATOMIC);
  var value_cell2 = Blockly.Python.valueToCode(block, 'Cell2', Blockly.Python.ORDER_ATOMIC);
  // TODO: Assemble Python into code variable.
  var code = '...';
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.Python.ORDER_NONE];
};

Blockly.Python['excel_set_value'] = function(block) {
  var value_rng = Blockly.Python.valueToCode(block, 'rng', Blockly.Python.ORDER_ATOMIC);
  var value_val = Blockly.Python.valueToCode(block, 'val', Blockly.Python.ORDER_ATOMIC);
  // TODO: Assemble Python into code variable.
  var code = '...\n';
  return code;
};

Blockly.Python['excel_get_value'] = function(block) {
  var value_rng = Blockly.Python.valueToCode(block, 'rng', Blockly.Python.ORDER_ATOMIC);
  // TODO: Assemble Python into code variable.
  var code = '...';
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.Python.ORDER_NONE];
};