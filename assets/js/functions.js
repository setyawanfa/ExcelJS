/**
 * Instantiate ExcelJS
 */
var ExcelJS = ExcelJS || {};

/**
 * Set ExcelJS Defaults
 */
ExcelJS.Table = function(){

    //Replicate Array.foreach for NodeList
    NodeList.prototype.forEach = Array.prototype.forEach;

    //Set Defaults
    this.exceljs    = document.getElementById('exceljs');
    this.thead      = this.exceljs.childNodes[0];
    this.tbody      = this.exceljs.childNodes[1];
    this.tfoot      = this.exceljs.childNodes[2];
    this.cols       = 0;
    this.rows       = 0;
};

/**
 * Adds a column to the ExcelJS table.
 */
ExcelJS.Table.prototype.addColumn = function () {

    //Set _self to use 'this' within event listeners.
    var _self = this;

    //If no rows exist, then this.addRow will add first row and column.
    if(this.rows === 0){
        this.addRow();
    }
    else
    {
        //Create new table head and foot elements.
        var rows        = this.tbody.childNodes,
            cols        = this.cols+parseInt(1),
            rowKey      = 0,
            newColHead  = document.createElement('th'),
            colHead     = document.createElement('span');
            newColFoot  = document.createElement('td'),
            colFoot     = document.createElement('a');

        //Add new table header cell.
        colHead.innerHTML       = "Col "+(cols+parseInt(1));
        newColHead.className    = "th-col-"+cols+" col-"+cols;
        newColHead.setAttribute('data-sort', 'desc');
        newColHead.setAttribute('data-col-id', cols);
        newColHead.appendChild(colHead);
        this.thead.appendChild(newColHead);

        //Add sort column listener to table header cell.
        var th = document.getElementsByClassName("th-col-"+cols);
        th[0].addEventListener("click", function() {
            _self.sortColumn(this.getAttribute('data-col-id'), this.getAttribute('data-sort'));
        }, false);

        //Add table footer cell
        newColFoot.className    = "col-"+cols;
        colFoot.innerHTML       = "Delete Col";
        colFoot.className       = "col-"+cols+" deleteCol deleteCol-"+cols;
        colFoot.setAttribute("data-col-id", cols);
        newColFoot.appendChild(colFoot);
        this.tfoot.appendChild(newColFoot);

        //Loop rows, adding inputs with correct identifier.
        rows.forEach(function(row){

            var newInput        = document.createElement('input'),
                newCol          = document.createElement('td'),
                colIdentifier   = "col_"+rowKey+"_"+cols;

            newInput.type       = "text";
            newInput.id         = colIdentifier;
            newInput.className  = colIdentifier;
            newInput.value      = colIdentifier;
            newInput.readOnly   = true;
            newInput.setAttribute('data-row', rowKey);
            row.insertBefore(newCol, row.lastChild).appendChild(newInput);

            //Add click/blur listener for readonly attribute.
            var field = document.getElementsByClassName("col_"+rowKey+"_"+cols);
            field[0].addEventListener("click", function() {
                this.readOnly   = false;
            }, false);
            field[0].addEventListener("blur", function() {
                this.readOnly   = true;
            }, false);

            rowKey++;
        });

        this.cols++;
    }

    if(typeof cols !== 'undefined'){
        //Add event listener to delete column link.
        var deleteCol = document.getElementsByClassName("deleteCol-"+cols);
        deleteCol[0].addEventListener("click", function() {
            _self.deleteColumn(this.getAttribute("data-col-id"))
            _self.cols--;
        }, false);
    }

    //Need to ensure all keys are correct for reference purposes.
    this.resetKeys();
};

/**
 * Adds a row to the ExcelJS table.
 */
ExcelJS.Table.prototype.addRow = function () {

    //Set _self to use 'this' within event listeners.
    var _self = this;

    //Create new row elements.
    var newColHead  = document.createElement('th'),
        colHead     = document.createElement('span'),
        newRow      = document.createElement('tr'),
        deleteRow   = document.createElement('a'),
        newColFoot  = document.createElement('td'),
        colFoot     = document.createElement('a');


    //If there are no rows, we need to add the first row and column.
    if(this.rows === 0){

        //Add Header
        colHead.innerHTML       = "Col "+(this.cols+parseInt(1));
        newColHead.className    = "th-col-"+this.cols+" col-"+this.cols;
        newColHead.setAttribute('data-sort', 'desc');
        newColHead.setAttribute('data-col-id', this.cols);
        newColHead.appendChild(colHead);
        this.thead.appendChild(newColHead);

        //Add header sort listener.
        var th = document.getElementsByClassName("th-col-"+this.cols);
        th[0].addEventListener("click", function() {
            _self.sortColumn(this.getAttribute('data-col-id'), this.getAttribute('data-sort'));
        }, false);

        //Add footer delete link.
        colFoot.innerHTML       = "Delete Col";
        colFoot.setAttribute("data-col-id", this.cols);
        colFoot.className       = "col-"+this.cols+" deleteCol deleteCol-"+this.cols;
        newColFoot.appendChild(colFoot);
        this.tfoot.appendChild(newColFoot);

        //Add eventlistener to delete column link.
        var deleteCol = document.getElementsByClassName("deleteCol-"+this.cols);
        deleteCol[0].addEventListener("click", function() {
            _self.deleteColumn(this.getAttribute("data-col-id"))
            _self.cols--;
        }, false);
    }

    newRow.id           = "row_"+this.rows;

    //Ensure this.cols is not negative.
    if(this.cols < 0){
        this.cols = 0;
    }

    //Loop columns and add correct number of columns to new row.
    var counter = 0;
    while(counter <= this.cols)
    {
        var newInput    = document.createElement('input');
        newInput.type       = "text";
        newInput.className  = "col_"+this.rows+"_"+counter;
        newInput.readOnly   = true;
        newInput.value      = "col_"+this.rows+"_"+counter;
        newInput.setAttribute('data-row', this.rows);
        newRow.insertCell().appendChild(newInput);
        counter++;
    }

    //Add delete row link as the row last child.
    deleteRow.innerHTML = "Delete Row";
    deleteRow.className = "deleteRow";
    deleteRow.setAttribute("data-row-id", this.rows);
    newRow.insertCell().appendChild(deleteRow);
    this.tbody.appendChild(newRow);

    //Now fields have been added, add event listener for changing readonly on click/blur.
    var counter = 0;
    while(counter <= this.cols)
    {
        var field = document.getElementsByClassName("col_"+this.rows+"_"+counter);
        field[0].addEventListener("click", function() {
            this.readOnly   = false;
        }, false);
        field[0].addEventListener("blur", function() {
            this.readOnly   = true;
        }, false);

        counter++;
    }

    //Add eventlistener to delete row button.
    var deleteRow = document.getElementsByClassName("deleteRow");
    deleteRow[this.rows].addEventListener("click", function() {
        var rowKey = this.getAttribute('data-row-id');
        _self.deleteRow(rowKey);
    }, false);

    //Need to ensure all keys are correct for reference purposes.
    this.resetKeys();

    //Row has been added, increment global counter.
    this.rows++;
};

/**
* Deletes row from the ExcelJS table. Called by delete row link event listener.
* @param  int rowKey This is the key of the row to be deleted.
*/
ExcelJS.Table.prototype.deleteRow = function (rowKey) {

    //Gets the row to be deleted, and removes from tbody.
    var row = document.getElementById("row_"+rowKey);
    row.remove();

    //Need to ensure all keys are correct for reference purposes.
    this.resetKeys();

    //Row has been removed, decrement global counter.
    this.rows--;

    //If there are now no rows, we don't need:
    //Header Sort Links, Delete Row Links, Delete Column Links
    //This means the thead and tfoot can also be cleared.
    if(this.rows === 0)
    {
        this.cols = 0;
        this.thead.innerHTML = '';
        this.tfoot.innerHTML = '';
    }

    //Since elements have been removed, we need to reset this.tbody.
    this.tbody = this.exceljs.childNodes[1];
};

/**
* Deletes column from the ExcelJS table. Called by delete column link event listener.
* @param  int columnKey This is the key of the column to be deleted.
*/
ExcelJS.Table.prototype.deleteColumn = function (columnKey) {

    //Gets the column elements to be deleted, and removes from tbody.
    var cols = document.getElementsByClassName("col-"+columnKey);
    while(cols.length > 0){
        cols[0].parentNode.removeChild(cols[0]);
    }

    //If there are no field columns then we need to clear the tbody and
    //reset the global row counter.
    if(this.cols === 0)
    {
        this.tbody.innerHTML = '';
        this.rows = 0;
    }

    //Need to ensure all keys are correct for reference purposes.
    this.resetKeys();
};

/**
* Deletes column from the ExcelJS table. Called by delete column link event listener.
* @param  int       columnKey   This is the key of the column to be deleted.
* @param  varchar   sort        Sort order [asc/desc].
*/
ExcelJS.Table.prototype.sortColumn = function (columnKey, sort) {

    //Set _self to use 'this' within event listeners.
    var _self = this;

    //Update th sort attribute for next click.
    var newSort = (sort === 'asc' ? 'desc' : 'asc'),
        th      = document.getElementsByClassName("th-col-"+columnKey);

    th[0].setAttribute('data-sort', newSort);

    //Get data ready for looping.
    var data            = this.tbody.getElementsByClassName("col-"+columnKey),
        sortCounter     = 0,
        inputValArray   = [];

    //Add input values to array for sorting.
    while(sortCounter < data.length)
    {
        var key = data.item(sortCounter).childNodes[0].getAttribute('data-row'),
            val = data.item(sortCounter).childNodes[0].value+"-"+key;

        inputValArray.push(val);
        sortCounter++;
    }

    //Sort the array alphanumerically ascending.
    inputValArray.sort();

    //If we are sorting descending, run reverse function also.
    if(sort === 'desc'){
        inputValArray.sort().reverse();
    }

    //Create a new tbody to replace old with new row order.
    var newTbody    = document.createElement('tbody');

    //Add rows to new tbody based on new row order.
    inputValArray.forEach(function(val){
        var rowKey  = val.split("-")[1];
        if(typeof rowKey !== 'undefined'){
            var row = document.getElementById('row_'+rowKey);
            newTbody.appendChild(row);
        }
    });

    //Replace tbody with new tbody containing new row order.
    this.exceljs.replaceChild(newTbody, this.exceljs.childNodes[1]);

    //Since tbody has been replaced, we need to reset this.tbody.
    this.tbody = this.exceljs.childNodes[1];
};

/**
 * Resets row and col keys when rows are added/deleted.
 */
ExcelJS.Table.prototype.resetKeys = function () {

    //Update thead cell keys.
    var row     = this.thead,
        cols    = row.childNodes,
        colKey  = 0;

    cols.forEach(function(col){
        col.className   = "th-col-"+colKey+" col-"+colKey;
        col.setAttribute('data-sort', 'desc');
        col.setAttribute('data-col-id', colKey);
        col.childNodes[0].innerHTML = "Col "+(colKey+parseInt(1));
        colKey++;
    });

    //Update tbody cell and input keys.
    var rows    = this.tbody.childNodes,
        rowKey  = 0;

    //Loop rows.
    rows.forEach(function(row){

        //Update row id.
        row.id = "row_"+rowKey;

        //Get childNodes and loop these to update keys.
        var cols = row.childNodes,
            colKey = 0;

        cols.forEach(function(col){
            //Update cell keys.
            col.id          = "col_"+rowKey+"_"+colKey;
            col.className   = "col-"+colKey+" col_"+rowKey+"_"+colKey;
            col.setAttribute('data-row', rowKey);

            //If the cell isn't the delete row cell, then update the input keys.
            if(col.firstChild.className !== 'deleteRow')
            {
                col.firstChild.className = "col_"+rowKey+"_"+colKey;
                col.firstChild.setAttribute('data-row', rowKey);
            }

            colKey++;
        });

        //Update delete row attribute for the row id key.
        var deleteRowColumn = row.lastChild.firstChild;
        deleteRowColumn.setAttribute("data-row-id", rowKey);

        rowKey++;
    });

    //Update tfoot cell keys.
    var row     = this.tfoot,
        cols    = row.childNodes,
        colKey  = 0;

    cols.forEach(function(col){
        col.className   = "col-"+colKey;
        col.childNodes[0].innerHTML = "Delete Col";
        col.childNodes[0].className = "col-"+colKey+" deleteCol deleteCol-"+colKey;
        col.childNodes[0].setAttribute("data-col-id", colKey);
        colKey++;
    });

    //Since elements have been changed, we need to reset this.tbody.
    this.tbody = this.exceljs.childNodes[1];
};

(function (ExcelJS) {

    //Instantiate Table as new ExcelJS
    var Table = new ExcelJS.Table();

    //Add new column listener.
    document.getElementById("addColumn").addEventListener("click", function(e) {
    	e.preventDefault();
        Table.addColumn();
    });
    //Add new row listener.
    document.getElementById("addRow").addEventListener("click", function(e) {
    	e.preventDefault();
        Table.addRow();
    });

}(ExcelJS));
