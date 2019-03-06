Blockly.Blocks['sheet_init'] = {
    init: function () {
      this.appendValueInput("sheetUrl")
        .appendField("載入 Google 試算表網址")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT);
      this.appendValueInput("sheetName")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("工作表名稱");
      this.setPreviousStatement(true, null);
      this.setNextStatement(true, null);
      this.setColour(200);
      this.setTooltip("");
      this.setHelpUrl("");
    }
  };
  
  Blockly.Blocks['sheet_write_normal'] = {
    init: function () {
      this.appendValueInput("range")
        .setCheck(null)
        .setCheck("String")
        .appendField("從 Google 試算表的");
      this.appendValueInput("data")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("寫入資料");
      this.setInputsInline(true);
      this.setPreviousStatement(true, null);
      this.setNextStatement(true, null);
      this.setColour(200);
      this.setTooltip("");
      this.setHelpUrl("");
      this.setCommentText("試算表寫入資料之後，才會繼續執行下方程式");
    }
  };
  
  
  Blockly.Blocks['sheet_write_easy'] = {
    init: function () {
      this.appendDummyInput()
        .appendField("從 Google 試算表的")
        .appendField(new Blockly.FieldDropdown([["最上方", "first"], ["最下方", "last"]]), "type")
        .appendField("寫入資料");
      this.appendValueInput('data_0')
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField('欄位 A 值:');
      this.appendValueInput('data_1')
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField('欄位 B 值:');
      this.setPreviousStatement(true);
      this.setNextStatement(true);
      this.setMutator(new Blockly.Mutator(['sheet_write_easy_item']));
      this.setColour(200);
      this.setTooltip("");
      this.setHelpUrl("");
      this.setCommentText("試算表寫入資料之後，才會繼續執行下方程式");
      this.itemCount_ = 2;
    },
    mutationToDom: function (workspace) {
      var container = document.createElement('mutation');
      container.setAttribute('items', this.itemCount_);
      return container;
    },
    domToMutation: function (container) {
      var Alphabet = ['0', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
      for (var x = 0; x < this.itemCount_; x++) {
        this.removeInput('data_' + x);
      }
      this.itemCount_ = parseInt(container.getAttribute('items'), 10);
      for (var x = 0; x < this.itemCount_; x++) {
        if (x < 27) {
          var input = this.appendValueInput('data_' + x)
            .setAlign(Blockly.ALIGN_RIGHT)
            .appendField(Blockly.Msg.WEBDUINO_GOOGLESHEETS_COLUMNWITE + Alphabet[x + 1] + Blockly.Msg.WEBDUINO_GOOGLESHEETS_COLUMNVALUE);
        } else {
          var input = this.appendValueInput('data_' + x)
            .setAlign(Blockly.ALIGN_RIGHT)
            .appendField(Blockly.Msg.WEBDUINO_GOOGLESHEETS_COLUMNWITE + (x + 1) + Blockly.Msg.WEBDUINO_GOOGLESHEETS_COLUMNVALUE);
        }
      }
    },
    decompose: function (workspace) {
      var containerBlock = Blockly.Block.obtain(workspace, 'sheet_write_easy_container');
      containerBlock.initSvg();
      var connection = containerBlock.getInput('STACK').connection;
      for (var x = 0; x < this.itemCount_; x++) {
        var optionBlock = Blockly.Block.obtain(workspace, 'sheet_write_easy_item');
        optionBlock.initSvg();
        connection.connect(optionBlock.previousConnection);
        connection = optionBlock.nextConnection;
      }
      return containerBlock;
    },
    compose: function (containerBlock) {
      var Alphabet = ['0', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
      for (var x = this.itemCount_ - 1; x >= 0; x--) {
        this.removeInput('data_' + x);
      }
      this.itemCount_ = 0;
      var optionBlock = containerBlock.getInputTargetBlock('STACK');
      while (optionBlock) {
        if (this.itemCount_ < 27) {
          var input = this.appendValueInput('data_' + this.itemCount_)
            .setAlign(Blockly.ALIGN_RIGHT)
            .appendField(Blockly.Msg.WEBDUINO_GOOGLESHEETS_COLUMNWITE + Alphabet[this.itemCount_ + 1] + Blockly.Msg.WEBDUINO_GOOGLESHEETS_COLUMNVALUE);
        } else {
          var input = this.appendValueInput('data_' + this.itemCount_)
            .setAlign(Blockly.ALIGN_RIGHT)
            .appendField(Blockly.Msg.WEBDUINO_GOOGLESHEETS_COLUMNWITE + (this.itemCount_ + 1) + Blockly.Msg.WEBDUINO_GOOGLESHEETS_COLUMNVALUE);
        }
        if (optionBlock.dataData_) {
          input.connection.connect(optionBlock.dataData_);
        }
        this.itemCount_++;
        optionBlock = optionBlock.nextConnection &&
          optionBlock.nextConnection.targetBlock();
      }
    },
    saveConnections: function (containerBlock) {
      var optionBlock = containerBlock.getInputTargetBlock('STACK');
      var x = 0;
      while (optionBlock) {
        var name = this.getFieldValue('name_' + x);
        var data = this.getInput('data_' + x);
        optionBlock.nameData_ = name;
        optionBlock.dataData_ = data && data.connection.targetConnection;
        x++;
        optionBlock = optionBlock.nextConnection &&
          optionBlock.nextConnection.targetBlock();
      }
    },
    newQuote_: Blockly.Blocks['text'].newQuote_
  };
  
  
  Blockly.Blocks['sheet_write_easy_container'] = {
    init: function () {
      this.setColour(100);
      this.appendDummyInput()
        .appendField('增加資料欄位');
      this.appendStatementInput('STACK');
      this.setTooltip('');
      this.setColour(200);
      this.contextMenu = false;
    }
  };
  
  Blockly.Blocks['sheet_write_easy_item'] = {
    init: function () {
      this.setColour(100);
      this.appendDummyInput()
        .appendField('欄位');
      this.setPreviousStatement(true);
      this.setNextStatement(true);
      this.setTooltip('');
      this.setColour(230);
      this.contextMenu = false;
    }
  };
  
  Blockly.Blocks['sheet_read'] = {
    init: function () {
      this.appendDummyInput()
        .appendField("從 Google 試算表讀取資料");
      this.setPreviousStatement(true, null);
      this.setNextStatement(true, null);
      this.setColour(200);
      this.setTooltip("");
      this.setHelpUrl("");
      this.setCommentText("讀取試算表所有的資料，讀取之後，才會繼續執行下方程式");
    }
  };
  
  
  Blockly.Blocks['sheet_read_data'] = {
    init: function () {
      this.appendValueInput("range")
        .appendField("儲存格")
        .setCheck(null);
      this.appendDummyInput()
        .appendField("的資料");
      this.setOutput(true, null);
      this.setColour(230);
      this.setTooltip("");
      this.setHelpUrl("");
    }
  };
  
  Blockly.Blocks['sheet_read_data_all'] = {
    init: function () {
      this.appendDummyInput()
        .appendField("所有的資料");
      this.setOutput(true, null);
      this.setColour(230);
      this.setTooltip("");
      this.setHelpUrl("");
    }
  };
  
  Blockly.Blocks['sheet_read_data_last'] = {
    init: function () {
      this.appendDummyInput()
        .appendField("最後一")
        .appendField(new Blockly.FieldDropdown([["列", "lastRow"], ["欄", "lastColumn"]]), "type")
        .appendField("的號碼");
      this.setInputsInline(true);
      this.setOutput(true, null);
      this.setColour(230);
      this.setTooltip("");
      this.setHelpUrl("");
    }
  };
  
  Blockly.Blocks['sheet_delete_row'] = {
    init: function () {
      this.appendValueInput("num")
        .setCheck(null)
        .appendField("從 Google 試算表刪除第");
      this.appendValueInput("delete_num")
        .setCheck(null)
        .appendField("列到第");
      this.appendDummyInput()
        .appendField("列");
      this.setInputsInline(true);
      this.setPreviousStatement(true, null);
      this.setNextStatement(true, null);
      this.setColour(200);
      this.setTooltip("");
      this.setHelpUrl("");
      this.setCommentText("刪除之後，才會繼續執行下方程式");
    }
  };
  
  Blockly.Blocks['sheet_delete_column'] = {
    init: function () {
      this.appendValueInput("num")
        .setCheck(null)
        .appendField("從 Google 試算表刪除第");
      this.appendValueInput("delete_num")
        .setCheck(null)
        .appendField("欄到第");
      this.appendDummyInput()
        .appendField("欄");
      this.setInputsInline(true);
      this.setPreviousStatement(true, null);
      this.setNextStatement(true, null);
      this.setColour(200);
      this.setTooltip("");
      this.setHelpUrl("");
      this.setCommentText("刪除之後，才會繼續執行下方程式");
    }
  };
  
  Blockly.Blocks['sheet_add_row'] = {
    init: function () {
      this.appendValueInput("num")
        .setCheck(null)
        .appendField("從 Google 試算表的第");
      this.appendValueInput("add_num")
        .setCheck(null)
        .appendField("列的")
        .appendField(new Blockly.FieldDropdown([["上面", "above"], ["下面", "after"]]), "add_type")
        .appendField("增加");
      this.appendDummyInput()
        .appendField("列");
      this.setPreviousStatement(true, null);
      this.setNextStatement(true, null);
      this.setColour(200);
      this.setTooltip("");
      this.setHelpUrl("");
      this.setCommentText("增加之後，才會繼續執行下方程式");
    }
  };
  
  Blockly.Blocks['sheet_add_column'] = {
    init: function () {
      this.appendValueInput("num")
        .setCheck(null)
        .appendField("從 Google 試算表的第");
      this.appendValueInput("add_num")
        .setCheck(null)
        .appendField("欄的")
        .appendField(new Blockly.FieldDropdown([["左邊", "above"], ["右邊", "after"]]), "add_type")
        .appendField("增加");
      this.appendDummyInput()
        .appendField("欄");
      this.setPreviousStatement(true, null);
      this.setNextStatement(true, null);
      this.setColour(200);
      this.setTooltip("");
      this.setHelpUrl("");
      this.setCommentText("增加之後，才會繼續執行下方程式");
    }
  };