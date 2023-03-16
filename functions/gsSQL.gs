// @author Chris Demmings - https://demmings.github.io/
/**
 * Query any sheet range using standard SQL SELECT syntax.
 * EXAMPLE :  gsSQL("select * from expenses where type = ?1", "expenses", A1:B, true, "travel")
 * 
 * @param {String} statement - SQL string 
 * @param {...any} parms - "table name",  SheetRange, [..."table name", SheetRange], OutputTitles (true/false), [...Bind Variable] 
 * @returns {any[][]} - Double array of selected data.  First index ROW, Second index COLUMN.
 * @customfunction
 */
function gsSQL(statement, ...parms) {     //  skipcq: JS-0128
    return GasSql.execute(statement, parms);
}

class GasSql {
    static execute(statement, parms) {
        if (parms.length === 0 || (parms.length > 0 && (Array.isArray(parms[0]) || parms[0] === ''))) {
            return GasSql.executeSqlv1(statement, parms);
        }
        else if (parms.length > 0 && typeof parms[0] === 'string') {
            return GasSql.executeSqlv2(statement, parms);
        }
        else {
            throw new Error("Invalid gsSQL() parameter list.");
        }
    }

    static executeSqlv1(statement, parms) {
        const sqlCmd = new Sql();
        let columnTitle = true;
        const bindings = [];

        //  If first item of parms is an array, the parms are assumed to be:
        // @param {any[][]} tableArr - {{"tableName", "sheetRange", cacheSeconds, hasColumnTitle}; {"name","range",cache,true};...}"
        // @param {Boolean} columnTitle - TRUE will add column title to output (default=TRUE)
        // @param {...any} bindings - Bind variables to match '?' in SQL statement.
        const tableArr = parms.length > 0 ? parms[0] : [];

        const tableList = GasSql.parseTableSettings(tableArr, statement);
        Logger.log(`gsSQL: tableList=${tableList}.  Statement=${statement}. List Len=${tableList.length}`);

        for (const tableDef of tableList) {
            sqlCmd.addTableData(tableDef[0], tableDef[1], tableDef[2], tableDef[3]);
        }
        columnTitle = parms.length > 1 ? parms[1] : true;

        for (let i = 2; i < parms.length; i++) {
            bindings.push(parms[i]);
        }

        sqlCmd.enableColumnTitle(columnTitle);

        for (const bind of bindings) {
            sqlCmd.addBindParameter(bind);
        }

        return sqlCmd.execute(statement);
    }

    static executeSqlv2(statement, parms) {
        const sqlCmd = new Sql();
        let columnTitle = true;
        const bindings = [];

        //  We expect:  "tableName", tableData[], ...["tableName", tableData[]], includeColumnOutput, ...bindings
        let i = 0;
        while (i + 1 < parms.length && typeof parms[i] !== 'boolean') {
            Logger.log(`Add Table: ${parms[i]}. Items=${parms[i + 1].length}`);
            sqlCmd.addTableData(parms[i], parms[i + 1], 0, true);
            i += 2;
        }
        if (i < parms.length && typeof parms[i] === 'boolean') {
            columnTitle = parms[i];
            i++
        }
        Logger.log(`Column Titles: ${columnTitle}`);
        while (i < parms.length) {
            Logger.log(`Add BIND Variable: ${parms[i]}`);
            bindings.push(parms[i]);
            i++
        }

        sqlCmd.enableColumnTitle(columnTitle);

        for (const bind of bindings) {
            sqlCmd.addBindParameter(bind);
        }

        return sqlCmd.execute(statement);
    }

    /**
     * 
     * @param {any[][]} tableArr - Referenced Table list.  This is normally the second parameter in gsSQL() custom function.  
     * It is a double array with first index for TABLE, and the second index are settings in the table. 
     * The setting index for each table is as follows:
     * * 0 - Table Name.
     * * 1 - Sheet Range.
     * * 2 - Cache seconds.
     * * 3 - First row contains title (for field name)
     * @param {String} statement - SQL SELECT statement.  If no data specified in 'tableArr', the SELECT is 
     * parsed and each referenced table is assumed to be a TAB name on the sheet.
     * @param {Boolean} randomOrder - Returned table list is randomized.
     * @returns {any[][]} - Data from 'tableArr' PLUS any extracted tables referenced from SELECT statement.
     * It is a double array with first index for TABLE, and the second index are settings in the table. 
     * The setting index for each table is as follows:
     * * 0 - Table Name.
     * * 1 - Sheet Range.
     * * 2 - Cache seconds.
     * * 3 - First row contains title (for field name)
     */
    static parseTableSettings(tableArr, statement = "", randomOrder = true) {
        let tableList = [];
        let referencedTableSettings = tableArr;

        //  Get table names from the SELECT statement when no table range info is given.
        if (tableArr.length === 0 && statement !== "") {
            referencedTableSettings = Sql.getReferencedTableNames(statement);
        }

        if (referencedTableSettings.length === 0) {
            throw new Error('Missing table definition {{"name","range",cache};{...}}');
        }

        Logger.log(`tableArr = ${referencedTableSettings}`);
        for (/** @type {any[]} */ const table of referencedTableSettings) {
            if (table.length === 1)
                table.push(table[0]);   // if NO RANGE, assumes table name is sheet name.
            if (table.length === 2)
                table.push(60);      //  default 0 second cache.
            if (table.length === 3)
                table.push(true);    //  default HAS column title row.
            if (table[1] === "")
                table[1] = table[0];    //  If empty range, assumes TABLE NAME is the SHEET NAME and loads entire sheet.
            if (table.length !== 4)
                throw new Error("Invalid table definition [name,range,cache,hasTitle]");

            tableList.push(table);
        }

        //  If called at the same time, loading similar tables in similar order - all processes
        //  just wait for table - but if loaded in different order, each process could be loading something.
        if (randomOrder)
            tableList = tableList.sort(() => Math.random() - 0.5);

        return tableList;
    }
}

/** Perform SQL SELECT using this class. */
class Sql {
    constructor() {
        /** @property {Map<String,Table>} - Map of referenced tables.*/
        this.tables = new Map();
        /** @property {Boolean} - Are column tables to be ouptout? */
        this.columnTitle = false;
        /** @property {BindData} - List of BIND data linked to '?' in statement. */
        this.bindData = new BindData();
        /** @property {String} - derived table name to output in column title replacing source table name. */
        this.columnTableNameReplacement = null;
    }

    /**
     * Add data for each referenced table in SELECT, before EXECUTE().
     * @param {String} tableName - Name of table referenced in SELECT.
     * @param {any} tableData - Either double array or a named range.
     * @param {Number} cacheSeconds - How long should loaded data be cached (default=0)
     * @param {Boolean} hasColumnTitle - Is first data row the column title?
     * @returns {Sql}
     */
    addTableData(tableName, tableData, cacheSeconds = 0, hasColumnTitle = true) {
        let tableInfo = null;

        if (Array.isArray(tableData)) {
            tableInfo = new Table(tableName)
                .setHasColumnTitle(hasColumnTitle)
                .loadArrayData(tableData);
        }
        else {
            tableInfo = new Table(tableName)
                .setHasColumnTitle(hasColumnTitle)
                .loadNamedRangeData(tableData, cacheSeconds);
        }

        this.tables.set(tableName.toUpperCase(), tableInfo);

        return this;
    }

    /**
     * Copies the data from an external tableMap to this instance.  
     * It copies a reference to outside array data only.  
     * The schema would need to be re-loaded.
     * @param {Map<String,Table>} tableMap 
     */
    copyTableData(tableMap) {
        // @ts-ignore
        for (const tableName of tableMap.keys()) {
            const tableInfo = tableMap.get(tableName);
            this.addTableData(tableName, tableInfo.tableData);
        }

        return this;
    }

    /**
     * Include column headers in return data.
     * @param {Boolean} value - true will return column names in first row of return data.
     * @returns {Sql}
     */
    enableColumnTitle(value) {
        this.columnTitle = value;
        return this;
    }

    /**
     * Derived table data that requires the ALIAS table name in column title.
     * @param {String} replacementTableName - derived table name to replace original table name.  To disable, set to null.
     * @returns {Sql}
     */
    replaceColumnTableNameWith(replacementTableName) {
        this.columnTableNameReplacement = replacementTableName;
        return this;
    }

    /**
     * Query if this instance of Sql() will generate column titles.
     * @returns {Boolean}
     */
    areColumnTitlesOutput() {
        return this.columnTitle;
    }

    /**
     * Add a bind data value.  Must be added in order.  If bind data is a named range, use addBindNamedRangeParameter().
     * @param {any} value - literal data. 
     * @returns {Sql}
     */
    addBindParameter(value) {
        this.bindData.add(value);
        return this;
    }

    /**
     * List of bind data added so far.
     * @returns {any[]}
     */
    getBindData() {
        return this.bindData.getBindDataList();
    }

    /**
     * The BIND data is a sheet named range that will be read and used for bind data.
     * @param {String} value - Sheets Named Range for SINGLE CELL only.
     * @returns {Sql}
     */
    addBindNamedRangeParameter(value) {
        const namedValue = TableData.getValueCached(value, 30);
        this.bindData.add(namedValue);
        Logger.log(`BIND=${value} = ${namedValue}`);
        return this;
    }

    /**
     * Set all bind data at once using array.
     * @param {BindData} value - Bind data.
     * @returns {Sql}
     */
    setBindValues(value) {
        this.bindData = value;
        return this;
    }

    /**
     * Clears existing BIND data so Sql() instance can be used again with new bind parameters.
     * @returns {Sql}
     */
    clearBindParameters() {
        this.bindData.clear();
        return this;
    }

    /**
    * Parse SQL SELECT statement, performs SQL query and returns data ready for custom function return.
    * <br>Execute() can be called multiple times for different SELECT statements, provided that all required
    * table data was loaded in the constructor.  
    * Methods that would be used PRIOR to execute are:
    * <br>**enableColumnTitle()** - turn on/off column title in output
    * <br>**addBindParameter()** - If bind data is needed in select.  e.g. "select * from table where id = ?"
    * <br>**addTableData()** - At least ONE table needs to be added prior to execute. This tells **execute** where to find the data.
    * <br>**Example SELECT and RETURN Data**
    * ```js
    *   let stmt = "SELECT books.id, books.title, books.author_id " +
    *        "FROM books " +
    *        "WHERE books.author_id IN ('11','12') " +
    *        "ORDER BY books.title";
    *
    *    let data = new Sql()
    *        .addTableData("books", this.bookTable())
    *        .enableColumnTitle(true)
    *        .execute(stmt);
    * 
    *    Logger.log(data);
    * 
    * [["books.id", "books.title", "books.author_id"],
    *    ["4", "Dream Your Life", "11"],
    *    ["8", "My Last Book", "11"],
    *    ["5", "Oranges", "12"],
    *    ["1", "Time to Grow Up!", "11"]]
    * ```
    * @param {any} statement - SELECT statement as STRING or AST of SELECT statement.
    * @returns {any[][]} - Double array where first index is ROW and second index is COLUMN.
    */
    execute(statement) {
        let sqlData = [];

        this.ast = (typeof statement === 'string') ? SqlParse.sql2ast(statement) : statement;

        //  "SELECT * from (select a,b,c from table) as derivedtable"
        //  Sub query data is loaded and given the name 'derivedtable' (using ALIAS from AS)
        //  The AST.FROM is updated from the sub-query to the new derived table name. 
        this.selectFromSubQuery();

        //  A JOIN table can a sub-query.  When this is the case, the sub-query SELECT is
        //  evaluated and the return data is given the ALIAS (as) name.  The AST is then
        //  updated to use the new table.
        this.selectJoinSubQuery();

        Sql.setTableAlias(this.tables, this.ast);
        Sql.loadSchema(this.tables);

        if (typeof this.ast.SELECT !== 'undefined') {
            sqlData = this.select(this.ast);
        }
        else
            throw new Error("Only SELECT statements are supported.");

        return sqlData;
    }

    /**
     * Updates 'tables' with table column information.
     * @param {Map<String,Table>} tables 
     */
    static loadSchema(tables) {
        // @ts-ignore
        for (const table of tables.keys()) {
            const tableInfo = tables.get(table.toUpperCase());
            tableInfo.loadSchema();
        }
    }

    /**
     * Updates 'tables' with associated table ALIAS name found in ast.
     * @param {Map<String,Table>} tables 
     * @param {Object} ast 
     */
    static setTableAlias(tables, ast) {
        // @ts-ignore
        for (const table of tables.keys()) {
            const tableAlias = Sql.getTableAlias(table, ast);
            const tableInfo = tables.get(table.toUpperCase());
            tableInfo.setTableAlias(tableAlias);
        }
    }

    /**
     * Sets all tables referenced SELECT.
    * @param {Map<String,Table>} mapOfTables - Map of referenced tables indexed by TABLE name.
    */
    setTables(mapOfTables) {
        this.tables = mapOfTables;
        return this;
    }

    /**
     * Returns a map of all tables configured for this SELECT.
     * @returns {Map<String,Table>} - Map of referenced tables indexed by TABLE name.
     */
    getTables() {
        return this.tables;
    }

    /**
    * Find table alias name (if any) for input actual table name.
    * @param {String} tableName - Actual table name.
    * @param {Object} ast - Abstract Syntax Tree for SQL.
    * @returns {String} - Table alias.  Empty string if not found.
    */
    static getTableAlias(tableName, ast) {
        let tableAlias = "";
        const ucTableName = tableName.toUpperCase();

        tableAlias = Sql.getTableAliasFromJoin(tableAlias, ucTableName, ast);
        tableAlias = Sql.getTableAliasUnion(tableAlias, ucTableName, ast);
        tableAlias = Sql.getTableAliasWhereIn(tableAlias, ucTableName, ast);
        tableAlias = Sql.getTableAliasWhereTerms(tableAlias, ucTableName, ast);

        return tableAlias;
    }

    /**
     * Modifies AST when FROM is a sub-query rather than a table name.
     */
    selectFromSubQuery() {
        if (typeof this.ast.FROM !== 'undefined' && typeof this.ast.FROM.SELECT !== 'undefined') {
            const data = new Sql()
                .setTables(this.tables)
                .enableColumnTitle(true)
                .replaceColumnTableNameWith(this.ast.FROM.table)
                .execute(this.ast.FROM);

            if (typeof this.ast.FROM.table !== 'undefined') {
                this.addTableData(this.ast.FROM.table, data);
            }

            if (this.ast.FROM.table === '') {
                throw new Error("Every derived table must have its own alias");
            }

            this.ast.FROM.as = '';
        }
    }

    /**
     * Checks if the JOINed table is a sub-query.  
     * The sub-query is evaluated and assigned the alias name.
     * The AST is adjusted to use the new JOIN TABLE.
     * @returns {void}
     */
    selectJoinSubQuery() {
        if (typeof this.ast.JOIN !== 'undefined') {
            for (const joinAst of this.ast.JOIN) {
                if (typeof joinAst.table !== 'string') {
                    const data = new Sql()
                        .setTables(this.tables)
                        .enableColumnTitle(true)
                        .replaceColumnTableNameWith(joinAst.as)
                        .execute(joinAst.table);

                    if (typeof joinAst.as !== 'undefined') {
                        this.addTableData(joinAst.as, data);
                    }

                    if (joinAst.as === '') {
                        throw new Error("Every derived table must have its own alias");
                    }
                    joinAst.table = joinAst.as;
                    joinAst.as = '';
                }
            }
        }
    }

    /**
     * Searches the FROM and JOIN components of a SELECT to find the table alias.
     * @param {String} tableAlias - Default alias name
     * @param {String} tableName - table name to search for.
     * @param {Object} ast - Abstract Syntax Tree to search
     * @returns {String} - Table alias name.
     */
    static getTableAliasFromJoin(tableAlias, tableName, ast) {
        const astTableBlocks = ['FROM', 'JOIN'];
        let aliasNameFound = tableAlias;

        let i = 0;
        while (aliasNameFound === "" && i < astTableBlocks.length) {
            aliasNameFound = Sql.locateAstTableAlias(tableName, ast, astTableBlocks[i]);
            i++;
        }

        return aliasNameFound;
    }

    /**
     * Searches the UNION portion of the SELECT to locate the table alias.
     * @param {String} tableAlias - default table alias.
     * @param {String} tableName - table name to search for.
     * @param {Object} ast - Abstract Syntax Tree to search
     * @returns {String} - table alias
     */
    static getTableAliasUnion(tableAlias, tableName, ast) {
        const astRecursiveTableBlocks = ['UNION', 'UNION ALL', 'INTERSECT', 'EXCEPT'];
        let extractedAlias = tableAlias;

        let i = 0;
        while (extractedAlias === "" && i < astRecursiveTableBlocks.length) {
            if (typeof ast[astRecursiveTableBlocks[i]] !== 'undefined') {
                for (const unionAst of ast[astRecursiveTableBlocks[i]]) {
                    extractedAlias = Sql.getTableAlias(tableName, unionAst);

                    if (extractedAlias !== "")
                        break;
                }
            }
            i++;
        }

        return extractedAlias;
    }

    /**
     * Search WHERE IN component of SELECT to find table alias.
     * @param {String} tableAlias - default table alias
     * @param {String} tableName - table name to search for
     * @param {Object} ast - Abstract Syntax Tree to search
     * @returns {String} - table alias
     */
    static getTableAliasWhereIn(tableAlias, tableName, ast) {
        let extractedAlias = tableAlias;
        if (tableAlias === "" && typeof ast.WHERE !== 'undefined' && ast.WHERE.operator === "IN") {
            extractedAlias = Sql.getTableAlias(tableName, ast.WHERE.right);
        }

        if (extractedAlias === "" && ast.operator === "IN") {
            extractedAlias = Sql.getTableAlias(tableName, ast.right);
        }

        return extractedAlias;
    }

    /**
     * Search WHERE terms of SELECT to find table alias.
     * @param {String} tableAlias - default table alias
     * @param {String} tableName  - table name to search for.
     * @param {Object} ast - Abstract Syntax Tree to search.
     * @returns {String} - table alias
     */
    static getTableAliasWhereTerms(tableAlias, tableName, ast) {
        let extractedTableAlias = tableAlias;
        if (tableAlias === "" && typeof ast.WHERE !== 'undefined' && typeof ast.WHERE.terms !== 'undefined') {
            for (const term of ast.WHERE.terms) {
                if (extractedTableAlias === "")
                    extractedTableAlias = Sql.getTableAlias(tableName, term);
            }
        }

        return extractedTableAlias;
    }

    /**
     * Create table definition array from select string.
     * @param {String} statement - full sql select statement.
     * @returns {String[][]} - table definition array.
     */
    static getReferencedTableNames(statement) {
        const ast = SqlParse.sql2ast(statement);
        return this.getReferencedTableNamesFromAst(ast);
    }

    /**
     * Create table definition array from select AST.
     * @param {Object} ast - AST for SELECT. 
     * @returns {any[]} - table definition array.
     * * [0] - table name.
     * * [1] - sheet tab name
     * * [2] - cache seconds
     * * [3] - output column title flag
     */
    static getReferencedTableNamesFromAst(ast) {
        const DEFAULT_CACHE_SECONDS = 60;
        const DEFAULT_COLUMNS_OUTPUT = true;
        const tableSet = new Map();

        Sql.extractAstTables(ast, tableSet);

        const tableList = [];
        // @ts-ignore
        for (const key of tableSet.keys()) {
            const tableDef = [key, key, DEFAULT_CACHE_SECONDS, DEFAULT_COLUMNS_OUTPUT];

            tableList.push(tableDef);
        }

        return tableList;
    }

    /**
     * Search for all referenced tables in SELECT.
     * @param {Object} ast - AST for SELECT.
     * @param {Map<String,String>} tableSet  - Function updates this map of table names and alias name.
     */
    static extractAstTables(ast, tableSet) {
        Sql.getTableNamesFrom(ast, tableSet);
        Sql.getTableNamesJoin(ast, tableSet);
        Sql.getTableNamesUnion(ast, tableSet);
        Sql.getTableNamesWhereIn(ast, tableSet);
        Sql.getTableNamesWhereTerms(ast, tableSet);
        Sql.getTableNamesCorrelatedSelect(ast, tableSet);
    }

    /**
     * Search for referenced table in FROM or JOIN part of select.
     * @param {Object} ast - AST for SELECT.
     * @param {Map<String,String>} tableSet  - Function updates this map of table names and alias name.
     */
    static getTableNamesFrom(ast, tableSet) {
        let fromAst = ast.FROM;
        while (typeof fromAst !== 'undefined') {
            if (typeof fromAst.isDerived === 'undefined') {
                tableSet.set(fromAst.table.toUpperCase(), typeof fromAst.as === 'undefined' ? '' : fromAst.as.toUpperCase());
            }
            else {
                Sql.extractAstTables(fromAst.FROM, tableSet);
            }
            fromAst = fromAst.FROM;
        }
    }

    /**
    * Search for referenced table in FROM or JOIN part of select.
    * @param {Object} ast - AST for SELECT.
    * @param {Map<String,String>} tableSet  - Function updates this map of table names and alias name.
    */
    static getTableNamesJoin(ast, tableSet) {

        if (typeof ast.JOIN === 'undefined')
            return;

        for (const astItem of ast.JOIN) {
            if (typeof astItem.table === 'string') {
                tableSet.set(astItem.table.toUpperCase(), typeof astItem.as === 'undefined' ? '' : astItem.as.toUpperCase());
            }
            else {
                Sql.extractAstTables(astItem.table, tableSet);
            }
        }
    }

    /**
     * Check if input is iterable.
     * @param {any} input - Check this object to see if it can be iterated. 
     * @returns {Boolean} - true - can be iterated.  false - cannot be iterated.
     */
    static isIterable(input) {
        if (input === null || input === undefined) {
            return false
        }

        return typeof input[Symbol.iterator] === 'function'
    }

    /**
     * Searches for table names within SELECT (union, intersect, except) statements.
     * @param {Object} ast - AST for SELECT
     * @param {Map<String,String>} tableSet - Function updates this map of table names and alias name.
     */
    static getTableNamesUnion(ast, tableSet) {
        const astRecursiveTableBlocks = ['UNION', 'UNION ALL', 'INTERSECT', 'EXCEPT'];

        for (const block of astRecursiveTableBlocks) {
            if (typeof ast[block] !== 'undefined') {
                for (const unionAst of ast[block]) {
                    this.extractAstTables(unionAst, tableSet);
                }
            }
        }
    }

    /**
     * Searches for tables names within SELECT (in, exists) statements.
     * @param {Object} ast - AST for SELECT
     * @param {Map<String,String>} tableSet - Function updates this map of table names and alias name.
     */
    static getTableNamesWhereIn(ast, tableSet) {
        //  where IN ().
        const subQueryTerms = ["IN", "NOT IN", "EXISTS", "NOT EXISTS"]
        if (typeof ast.WHERE !== 'undefined' && (subQueryTerms.indexOf(ast.WHERE.operator) !== -1)) {
            this.extractAstTables(ast.WHERE.right, tableSet);
        }

        if (subQueryTerms.indexOf(ast.operator) !== -1) {
            this.extractAstTables(ast.right, tableSet);
        }
    }

    /**
     * Search WHERE to find referenced table names.
     * @param {Object} ast -  AST to search.
     * @param {Map<String,String>} tableSet - Function updates this map of table names and alias name.
     */
    static getTableNamesWhereTerms(ast, tableSet) {
        if (typeof ast.WHERE !== 'undefined' && typeof ast.WHERE.terms !== 'undefined') {
            for (const term of ast.WHERE.terms) {
                this.extractAstTables(term, tableSet);
            }
        }
    }

    /**
     * Search for table references in the WHERE condition.
     * @param {Object} ast -  AST to search.
     * @param {Map<String,String>} tableSet - Function updates this map of table names and alias name. 
     */
    static getTableNamesWhereCondition(ast, tableSet) {
        const lParts = typeof ast.left === 'string' ? ast.left.split(".") : [];
        if (lParts.length > 1) {
            tableSet.set(lParts[0].toUpperCase(), "");
        }
        const rParts = typeof ast.right === 'string' ? ast.right.split(".") : [];
        if (rParts.length > 1) {
            tableSet.set(rParts[0].toUpperCase(), "");
        }
        if (typeof ast.terms !== 'undefined') {
            for (const term of ast.terms) {
                Sql.getTableNamesWhereCondition(term, tableSet);
            }
        }
    }

    /**
     * Search CORRELATES sub-query for table names.
     * @param {*} ast - AST to search
     * @param {*} tableSet - Function updates this map of table names and alias name.
     */
    static getTableNamesCorrelatedSelect(ast, tableSet) {
        if (typeof ast.SELECT !== 'undefined') {
            for (const term of ast.SELECT) {
                if (typeof term.subQuery !== 'undefined' && term.subQuery !== null) {
                    this.extractAstTables(term.subQuery, tableSet);
                }
            }
        }
    }

    /**
     * Search a property of AST for table alias name.
     * @param {String} tableName - Table name to find in AST.
     * @param {Object} ast - AST of SELECT.
     * @param {String} astBlock - AST property to search.
     * @returns {String} - Alias name or "" if not found.
     */
    static locateAstTableAlias(tableName, ast, astBlock) {
        if (typeof ast[astBlock] === 'undefined')
            return "";

        let block = [ast[astBlock]];
        if (this.isIterable(ast[astBlock])) {
            block = ast[astBlock];
        }

        for (const astItem of block) {
            if (typeof astItem.table === 'string' && tableName === astItem.table.toUpperCase() && astItem.as !== "") {
                return astItem.as;
            }
        }

        return "";
    }

    /**
     * Load SELECT data and return in double array.
     * @param {Object} selectAst - Abstract Syntax Tree of SELECT
     * @returns {any[][]} - double array useable by Google Sheet in custom function return value.
     * * First row of data will be column name if column title output was requested.
     * * First Array Index - ROW
     * * Second Array Index - COLUMN
     */
    select(selectAst) {
        let recordIDs = [];
        let viewTableData = [];
        let ast = selectAst;

        if (typeof ast.FROM === 'undefined')
            throw new Error("Missing keyword FROM");

        //  Manipulate AST to add GROUP BY if DISTINCT keyword.
        ast = Sql.distinctField(ast);

        //  Manipulate AST add pivot fields.
        ast = this.pivotField(ast);

        const view = new SelectTables(ast, this.tables, this.bindData);

        //  JOIN tables to create a derived table.
        view.join(ast);                 // skipcq: JS-D008

        view.updateSelectedFields(ast);

        //  Get the record ID's of all records matching WHERE condition.
        recordIDs = view.whereCondition(ast);

        //  Get selected data records.
        viewTableData = view.getViewData(recordIDs);

        //  Compress the data.
        viewTableData = view.groupBy(ast, viewTableData);

        //  Sort our selected data.
        view.orderBy(ast, viewTableData);

        //  Remove fields referenced but not included in SELECT field list.
        view.removeTempColumns(viewTableData);

        if (typeof ast.LIMIT !== 'undefined') {
            const maxItems = ast.LIMIT.nb;
            if (viewTableData.length > maxItems)
                viewTableData.splice(maxItems);
        }

        //  Apply SET rules for various union types.
        viewTableData = this.unionSets(ast, viewTableData);

        if (this.columnTitle) {
            viewTableData.unshift(view.getColumnTitles(this.columnTableNameReplacement));
        }

        if (viewTableData.length === 0) {
            viewTableData.push([""]);
        }

        if (viewTableData.length === 1 && viewTableData[0].length === 0) {
            viewTableData[0] = [""];
        }

        return viewTableData;
    }

    /**
     * If 'GROUP BY' is not set and 'DISTINCT' column is specified, update AST to add 'GROUP BY'.
     * @param {Object} ast - Abstract Syntax Tree for SELECT.
     * @returns {Object} - Updated AST to include GROUP BY when DISTINCT field used.
     */
    static distinctField(ast) {
        const astFields = ast.SELECT;

        if (astFields.length > 0) {
            const firstField = astFields[0].name.toUpperCase();
            if (firstField.startsWith("DISTINCT")) {
                astFields[0].name = firstField.replace("DISTINCT", "").trim();

                if (typeof ast['GROUP BY'] === 'undefined') {
                    const groupBy = [];

                    for (const astItem of astFields) {
                        groupBy.push({ name: astItem.name, as: '' });
                    }

                    ast["GROUP BY"] = groupBy;
                }
            }
        }

        return ast;
    }

    /**
     * Add new column to AST for every AGGREGATE function and unique pivot column data.
     * @param {Object} ast - AST which is checked to see if a PIVOT is used.
     * @returns {Object} - Updated AST containing SELECT FIELDS for the pivot data OR original AST if no pivot.
     */
    pivotField(ast) {
        //  If we are doing a PIVOT, it then requires a GROUP BY.
        if (typeof ast.PIVOT !== 'undefined') {
            if (typeof ast['GROUP BY'] === 'undefined')
                throw new Error("PIVOT requires GROUP BY");
        }
        else
            return ast;

        // These are all of the unique PIVOT field data points.
        const pivotFieldData = this.getUniquePivotData(ast);

        ast.SELECT = Sql.addCalculatedPivotFieldsToAst(ast, pivotFieldData);

        return ast;
    }

    /**
     * Find distinct pivot column data.
     * @param {Object} ast - Abstract Syntax Tree containing the PIVOT option.
     * @returns {any[][]} - All unique data points found in the PIVOT field for the given SELECT.
     */
    getUniquePivotData(ast) {
        const pivotAST = {};

        pivotAST.SELECT = ast.PIVOT;
        pivotAST.SELECT[0].name = `DISTINCT ${pivotAST.SELECT[0].name}`;
        pivotAST.FROM = ast.FROM;
        pivotAST.WHERE = ast.WHERE;

        const pivotSql = new Sql()
            .enableColumnTitle(false)
            .setBindValues(this.bindData)
            .copyTableData(this.getTables());

        // These are all of the unique PIVOT field data points.
        const tableData = pivotSql.execute(pivotAST);

        return tableData;
    }

    /**
     * Add new calculated fields to the existing SELECT fields.  A field is add for each combination of
     * aggregate function and unqiue pivot data points.  The CASE function is used for each new field.
     * A test is made if the column data equal the pivot data.  If it is, the aggregate function data 
     * is returned, otherwise null.  The GROUP BY is later applied and the appropiate pivot data will
     * be calculated.
     * @param {Object} ast - AST to be updated.
     * @param {any[][]} pivotFieldData - Table data with unique pivot field data points. 
     * @returns {Object} - Abstract Sytax Tree with new SELECT fields with a CASE for each pivot data and aggregate function.
     */
    static addCalculatedPivotFieldsToAst(ast, pivotFieldData) {
        const newPivotAstFields = [];

        for (const selectField of ast.SELECT) {
            //  If this is an aggregrate function, we will add one for every pivotFieldData item
            const functionNameRegex = /^\w+\s*(?=\()/;
            const matches = selectField.name.match(functionNameRegex)
            if (matches !== null && matches.length > 0) {
                const args = SelectTables.parseForFunctions(selectField.name, matches[0].trim());

                for (const fld of pivotFieldData) {
                    const caseTxt = `${matches[0]}(CASE WHEN ${ast.PIVOT[0].name} = '${fld}' THEN ${args[1]} ELSE 'null' END)`;
                    const asField = `${fld[0]} ${typeof selectField.as !== 'undefined' && selectField.as !== "" ? selectField.as : selectField.name}`;
                    newPivotAstFields.push({ name: caseTxt, as: asField });
                }
            }
            else
                newPivotAstFields.push(selectField);
        }

        return newPivotAstFields;
    }

    /**
     * If any SET commands are found (like UNION, INTERSECT,...) the additional SELECT is done.  The new
     * data applies the SET rule against the income viewTableData, and the result data set is returned.
     * @param {Object} ast - SELECT AST.
     * @param {any[][]} viewTableData - SELECTED data before UNION.
     * @returns {any[][]} - New data with set rules applied.
     */
    unionSets(ast, viewTableData) {
        const unionTypes = ['UNION', 'UNION ALL', 'INTERSECT', 'EXCEPT'];
        let unionTableData = viewTableData;

        for (const type of unionTypes) {
            if (typeof ast[type] !== 'undefined') {
                const unionSQL = new Sql()
                    .setBindValues(this.bindData)
                    .copyTableData(this.getTables());
                for (const union of ast[type]) {
                    const unionData = unionSQL.execute(union);
                    if (unionTableData.length > 0 && unionData.length > 0 && unionTableData[0].length !== unionData[0].length)
                        throw new Error(`Invalid ${type}.  Selected field counts do not match.`);

                    switch (type) {
                        case "UNION":
                            //  Remove duplicates.
                            unionTableData = Sql.appendUniqueRows(unionTableData, unionData);
                            break;

                        case "UNION ALL":
                            //  Allow duplicates.
                            unionTableData = unionTableData.concat(unionData);
                            break;

                        case "INTERSECT":
                            //  Must exist in BOTH tables.
                            unionTableData = Sql.intersectRows(unionTableData, unionData);
                            break;

                        case "EXCEPT":
                            //  Remove from first table all rows that match in second table.
                            unionTableData = Sql.exceptRows(unionTableData, unionData);
                            break;

                        default:
                            throw new Error(`Internal error.  Unsupported UNION type: ${type}`);
                    }
                }
            }
        }

        return unionTableData;
    }

    /**
     * Appends any row in newData that does not exist in srcData.
     * @param {any[][]} srcData - existing table data
     * @param {any[][]} newData - new table data
     * @returns {any[][]} - srcData rows PLUS any row in newData that is NOT in srcData.
     */
    static appendUniqueRows(srcData, newData) {
        const srcMap = new Map();

        for (const srcRow of srcData) {
            srcMap.set(srcRow.join("::"), true);
        }

        for (const newRow of newData) {
            const key = newRow.join("::");
            if (!srcMap.has(key)) {
                srcData.push(newRow);
                srcMap.set(key, true);
            }
        }
        return srcData;
    }

    /**
     * Finds the rows that are common between srcData and newData
     * @param {any[][]} srcData - table data
     * @param {any[][]} newData - table data
     * @returns {any[][]} - returns only rows that intersect srcData and newData.
     */
    static intersectRows(srcData, newData) {
        const srcMap = new Map();
        const intersectTable = [];

        for (const srcRow of srcData) {
            srcMap.set(srcRow.join("::"), true);
        }

        for (const newRow of newData) {
            if (srcMap.has(newRow.join("::"))) {
                intersectTable.push(newRow);
            }
        }
        return intersectTable;
    }

    /**
     * Returns all rows in srcData MINUS any rows that match it from newData.
     * @param {any[][]} srcData - starting table
     * @param {any[][]} newData  - minus table (if it matches srcData row)
     * @returns {any[][]} - srcData MINUS newData
     */
    static exceptRows(srcData, newData) {
        const srcMap = new Map();
        let rowNum = 0;
        for (const srcRow of srcData) {
            srcMap.set(srcRow.join("::"), rowNum);
            rowNum++;
        }

        const removeRowNum = [];
        for (const newRow of newData) {
            const key = newRow.join("::");
            if (srcMap.has(key)) {
                removeRowNum.push(srcMap.get(key));
            }
        }

        removeRowNum.sort((a, b) => b - a);
        for (rowNum of removeRowNum) {
            srcData.splice(rowNum, 1);
        }

        return srcData;
    }
}

/**
 * Store and retrieve bind data for use in WHERE portion of SELECT statement.
 */
class BindData {
    constructor() {
        this.clear();
    }

    /**
     * Reset the bind data.
     */
    clear() {
        this.next = 1;
        this.bindMap = new Map();
        this.bindQueue = [];
    }

    /**
     * Add bind data 
     * @param {any} data - bind data
     * @returns {String} - bind variable name for reference in SQL.  e.g.  first data point would return '?1'.
     */
    add(data) {
        const key = `?${this.next.toString()}`;
        this.bindMap.set(key, data);
        this.bindQueue.push(data);

        this.next++;

        return key;
    }

    /**
     * Add a list of bind data points.
     * @param {any[]} bindList 
     */
    addList(bindList) {
        for (const data of bindList) {
            this.add(data);
        }
    }

    /**
     * Pull out a bind data entry.
     * @param {String} name - Get by name or get NEXT if empty.
     * @returns {any}
     */
    get(name = "") {
        return name === '' ? this.bindQueue.shift() : this.bindMap.get(name);
    }

    /**
     * Return the ordered list of bind data.
     * @returns {any[]} - Current list of bind data.
     */
    getBindDataList() {
        return this.bindQueue;
    }
}



/** Data and methods for each (logical) SQL table. */
class Table {       //  skipcq: JS-0128
    /**
     * 
     * @param {String} tableName - name of sql table.
     */
    constructor(tableName) {
        /** @property {String} - table name. */
        this.tableName = tableName.toUpperCase();

        /** @property {any[][]} - table data. */
        this.tableData = [];

        /** @property {Map<String, Map<String,Number[]>>} - table indexes*/
        this.indexes = new Map();

        /** @property {Boolean} */
        this.hasColumnTitle = true;

        /** @property {Schema} */
        this.schema = new Schema()
            .setTableName(tableName)
            .setTable(this);
    }

    /**
     * Set associated table alias name to object.
     * @param {String} tableAlias - table alias that may be used to prefix column names.
     * @returns {Table}
     */
    setTableAlias(tableAlias) {
        this.schema.setTableAlias(tableAlias);
        return this;
    }

    /**
     * Indicate if data contains a column title row.
     * @param {Boolean} hasTitle 
     * * true - first row of data will contain unique column names
     * * false - first row of data will contain data.  Column names are then referenced as letters (A, B, ...)
     * @returns {Table}
     */
    setHasColumnTitle(hasTitle) {
        this.hasColumnTitle = hasTitle;

        return this;
    }

    /**
     * Load sheets named range of data into table.
     * @param {String} namedRange - defines where data is located in sheets.
     * * sheet name - reads entire sheet from top left corner.
     * * named range - reads named range for data.
     * * A1 notation - range of data using normal sheets notation like 'A1:C10'.  This may also include the sheet name like 'stocks!A1:C100'.
     * @param {Number} cacheSeconds - How many seconds to cache data so we don't need to make time consuming
     * getValues() from sheets.  
     * @returns {Table}
     */
    loadNamedRangeData(namedRange, cacheSeconds = 0) {
        this.tableData = TableData.loadTableData(namedRange, cacheSeconds);

        if (!this.hasColumnTitle) {
            this.addColumnLetters(this.tableData);
        }

        Logger.log(`Load Data: Range=${namedRange}. Items=${this.tableData.length}`);
        this.loadSchema();

        return this;
    }

    /**
     * Read table data from a double array rather than from sheets.
     * @param {any[]} tableData - Loaded table data with first row titles included.
     * @returns {Table}
     */

    loadArrayData(tableData) {
        if (typeof tableData === 'undefined' || tableData.length === 0)
            return this;

        if (!this.hasColumnTitle) {
            this.addColumnLetters(tableData);
        }

        this.tableData = Table.removeEmptyRecordsAtEndOfTable(tableData);

        this.loadSchema();

        return this;
    }

    /**
     * It is common to have extra empty records loaded at end of table.
     * Remove those empty records at END of table only.
     * @param {any[][]} tableData 
     * @returns {any[][]}
     */
    static removeEmptyRecordsAtEndOfTable(tableData) {
        let blankLines = 0;
        for (let i = tableData.length-1; i > 0; i--) {
            if (tableData[i].join().replace(/,/g, "").length > 0)
                break;
            blankLines++;
        }

        return tableData.slice(0, tableData.length-blankLines);
    }

    /**
     * Internal function for updating the loaded data to include column names using letters, starting from 'A', 'B',...
     * @param {any[][]} tableData - table data that does not currently contain a first row with column names.
     * @returns {any[][]} - updated table data that includes a column title row.
     */
    addColumnLetters(tableData) {
        if (tableData.length === 0)
            return [[]];

        const newTitleRow = [];

        for (let i = 1; i <= tableData[0].length; i++) {
            newTitleRow.push(this.numberToSheetColumnLetter(i));
        }
        tableData.unshift(newTitleRow);

        return tableData;
    }

    /**
     * Find the sheet column letter name based on position.  
     * @param {Number} number - Returns the sheets column name.  
     * 1 = 'A'
     * 2 = 'B'
     * 26 = 'Z'
     * 27 = 'AA'
     * @returns {String} - the column letter.
     */
    numberToSheetColumnLetter(number) {
        const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        let result = ""

        let charIndex = number % alphabet.length
        let quotient = number / alphabet.length
        if (charIndex - 1 === -1) {
            charIndex = alphabet.length
            quotient--;
        }
        result = alphabet.charAt(charIndex - 1) + result;
        if (quotient >= 1) {
            result = this.numberToSheetColumnLetter(quotient) + result;
        }

        return result;
    }

    /**
     * Read loaded table data and updates internal list of column information
     * @returns {Table}
     */
    loadSchema() {
        this.schema
            .setTableData(this.tableData)
            .load();

        return this;
    }

    /**
     * Find column number using the field name.
     * @param {String} fieldName - Valid field name.
     * @returns {Number} - column offset number starting at zero.
     */
    getFieldColumn(fieldName) {
        return this.schema.getFieldColumn(fieldName);
    }

    /**
    * Get field column index (starts at 0) for field names.
    * @param {String[]} fieldNames - list of valid field names.
    * @returns {Number[]} - list of column offsets, starting at zero corresponding to the input list of names.
    */
    getFieldColumns(fieldNames) {
        return this.schema.getFieldColumns(fieldNames);
    }

    /**
     * Find all field data for this table (or the derived table)
     * @returns {VirtualField[]} - field column information list
     */
    getAllVirtualFields() {
        return this.schema.getAllVirtualFields();
    }

    /**
     * Returns a list of all possible field names that could be used in the SELECT.
     * @returns {String[]} - List of field names.
     */
    getAllFieldNames() {
        return this.schema.getAllFieldNames();
    }

    /**
     * Returns table field names that are prefixed with table name.
     * @returns {String[]} - field names
     */
    getAllExtendedNotationFieldNames() {
        return this.schema.getAllExtendedNotationFieldNames();
    }

    /**
     * Find number of columns in table.
     * @returns {Number} - column count.
     */
    getColumnCount() {
        const fields = this.getAllExtendedNotationFieldNames();
        return fields.length;
    }

    /**
     * Return range of records from table.
     * @param {Number} startRecord - 1 is first record
     * @param {Number} lastRecord - -1 for all. Last = RecordCount().    
     * @param {Number[]} fields - fields to include in output
     * @returns {any[][]} - subset table data.
     */
    getRecords(startRecord, lastRecord, fields) {
        const selectedRecords = [];

        let minStartRecord = startRecord;
        if (minStartRecord < 1) {
            minStartRecord = 1;
        }

        let maxLastRecord = lastRecord;
        if (maxLastRecord < 0) {
            maxLastRecord = this.tableData.length - 1;
        }

        for (let i = minStartRecord; i <= maxLastRecord && i < this.tableData.length; i++) {
            const row = [];

            for (const col of fields) {
                row.push(this.tableData[i][col]);
            }

            selectedRecords.push(row);
        }

        return selectedRecords;
    }

    /**
     * Create a logical table index on input field name.
     * The resulting map is stored with the table.
     * The Map<fieldDataItem, [rowNumbers]> is stored.
     * @param {String} fieldName - field name to index.
     */
    addIndex(fieldName) {
        const indexedFieldName = fieldName.trim().toUpperCase();
        /** @type {Map<String,Number[]>} */
        const fieldValuesMap = new Map();

        const fieldIndex = this.schema.getFieldColumn(indexedFieldName);
        for (let i = 1; i < this.tableData.length; i++) {
            let value = this.tableData[i][fieldIndex];
            if (value !== null) {
                value = value.toString();
            }

            if (value !== "") {
                let rowNumbers = [];
                if (fieldValuesMap.has(value))
                    rowNumbers = fieldValuesMap.get(value);

                rowNumbers.push(i);
                fieldValuesMap.set(value, rowNumbers);
            }
        }

        this.indexes.set(indexedFieldName, fieldValuesMap);
    }

    /**
     * Return all row ID's where FIELD = SEARCH VALUE.
     * @param {String} fieldName - table column name (must be upper case and trimmed)
     * @param {any} searchValue - value to search for in index
     * @returns {Number[]} - all matching row numbers.
     */
    search(fieldName, searchValue) {
        const rows = [];

        const fieldValuesMap = this.indexes.get(fieldName);
        if (fieldValuesMap.has(searchValue))
            return fieldValuesMap.get(searchValue);
        return rows;
    }

    /**
     * Append table data from 'concatTable' to the end of this tables existing data.
     * @param {Table} concatTable - Append 'concatTable' data to end of current table data.
     * @returns {void}
     */
    concat(concatTable) {
        const fieldsThisTable = this.schema.getAllFieldNames();
        const fieldColumns = concatTable.getFieldColumns(fieldsThisTable);
        const data = concatTable.getRecords(1, -1, fieldColumns);
        this.tableData = this.tableData.concat(data);
    }

}

/** Class contains information about each column in the SQL table. */
class Schema {
    constructor() {
        /** @property {String} - Table name. */
        this.tableName = "";

        /** @property {String} - Alias name of table. */
        this.tableAlias = "";

        /** @property {any[][]} - Table data double array. */
        this.tableData = [];

        /** @property {Table} - Link to table info object. */
        this.tableInfo = null;

        /** @property {Boolean} - Is this a derived table. */
        this.isDerivedTable = this.tableName === DERIVEDTABLE;

        /** @property {Map<String,Number>} - String=Field Name, Number=Column Number */
        this.fields = new Map();

        /** @property {VirtualFields} */
        this.virtualFields = new VirtualFields();
    }

    /**
     * Set table name in this object.
     * @param {String} tableName - Table name to remember.
     * @returns {Schema}
     */
    setTableName(tableName) {
        this.tableName = tableName.toUpperCase();
        return this;
    }

    /**
     * Associate the table alias to this object.
     * @param {String} tableAlias - table alias name
     * @returns {Schema}  
     */
    setTableAlias(tableAlias) {
        this.tableAlias = tableAlias.toUpperCase();
        return this;
    }

    /**
     * Associate table data with this object.
     * @param {any[][]} tableData - double array of table data.
     * @returns {Schema}
     */
    setTableData(tableData) {
        this.tableData = tableData;
        return this;
    }

    /**
     * Set the existing 'Table' info.
     * @param {Table} tableInfo - table object.
     * @returns {Schema}
     */
    setTable(tableInfo) {
        this.tableInfo = tableInfo;
        return this;
    }

    /**
     * Retrieve all field names for this table.
     * @returns {String[]} - List of field names.
     */
    getAllFieldNames() {
        /** @type {String[]} */
        const fieldNames = [];

        // @ts-ignore
        for (const key of this.fields.keys()) {
            if (key !== "*")
                fieldNames.push(key);
        }

        return fieldNames;
    }

    /**
     * All table fields names with 'TABLE.field_name'.
     * @returns {String[]} - list of all field names with table prefix.
     */
    getAllExtendedNotationFieldNames() {
        /** @type {String[]} */
        const fieldNames = [];

        // @ts-ignore
        for (const [key, value] of this.fields.entries()) {
            if (value !== null) {
                const fieldParts = key.split(".");
                if (typeof fieldNames[value] === 'undefined' ||
                    (fieldParts.length === 2 && (fieldParts[0] === this.tableName || this.isDerivedTable)))
                    fieldNames[value] = key;
            }
        }

        return fieldNames;
    }

    /**
     * Get a list of all virtual field data associated with this table.
     * @returns {VirtualField[]}
     */
    getAllVirtualFields() {
        return this.virtualFields.getAllVirtualFields();
    }

    /**
     * Get the column number for the specified field name.
     * @param {String} field - Field name to find column number for.
     * @returns {Number} - Column number.
     */
    getFieldColumn(field) {
        const cols = this.getFieldColumns([field]);
        return cols[0];
    }

    /**
    * Get field column index (starts at 0) for field names.
    * @param {String[]} fieldNames - find columns for specific fields in table.
    * @returns {Number[]} - column numbers for each specified field.
    */
    getFieldColumns(fieldNames) {
        /** @type {Number[]} */
        const fieldIndex = [];

        for (const field of fieldNames) {
            let i = -1;

            if (this.fields.has(field.trim().toUpperCase()))
                i = this.fields.get(field.trim().toUpperCase());

            fieldIndex.push(i);
        }

        return fieldIndex;
    }

    /**
     * The field name is found in TITLE row of sheet.  These column titles
     * are TRIMMED, UPPERCASE and SPACES removed (made to UNDERSCORE).
     * SQL statements MUST reference fields with spaces converted to underscore.
     * @returns {Schema}
     */
    load() {
        this.fields = new Map();
        this.virtualFields = new VirtualFields();

        if (this.tableData.length === 0)
            return this;

        /** @type {any[]} */
        const titleRow = this.tableData[0];

        let colNum = 0;
        /** @type {FieldVariants} */
        let fieldVariants = null;
        for (const baseColumnName of titleRow) {
            //  Find possible variations of the field column name.
            try {
                fieldVariants = this.getColumnNameVariants(baseColumnName);
            }
            catch (ex) {
                throw new Error(`Invalid column title: ${baseColumnName}`);
            }
            const columnName = fieldVariants.columnName;

            this.setFieldVariantsColumNumber(fieldVariants, colNum);

            if (columnName !== "") {
                const virtualField = new VirtualField(columnName, this.tableInfo, colNum);
                this.virtualFields.add(virtualField, true);
            }

            colNum++;
        }

        //  Add special field for every table.
        //  The asterisk represents ALL fields in table.
        this.fields.set("*", null);

        return this;
    }

    /**
     * @typedef {Object} FieldVariants
     * @property {String} columnName
     * @property {String} fullColumnName
     * @property {String} fullColumnAliasName
     */
    /**
     * Find all valid variations for a column name.  This will include base column name,
     * the column name prefixed with full table name, and the column name prefixed with table alias.
     * @param {String} colName 
     * @returns {FieldVariants}
     */
    getColumnNameVariants(colName) {
        const columnName = colName.trim().toUpperCase().replace(/\s/g, "_");
        let fullColumnName = columnName;
        let fullColumnAliasName = "";
        if (columnName.indexOf(".") === -1) {
            fullColumnName = `${this.tableName}.${columnName}`;
            if (this.tableAlias !== "")
                fullColumnAliasName = `${this.tableAlias}.${columnName}`;
        }

        return {columnName, fullColumnName, fullColumnAliasName};
    }

    /**
     * Associate table column number to each possible variation of column name.
     * @param {FieldVariants} fieldVariants 
     * @param {Number} colNum 
     */
    setFieldVariantsColumNumber(fieldVariants, colNum) {
        if (fieldVariants.columnName !== "") {
            this.fields.set(fieldVariants.columnName, colNum);

            if (!this.isDerivedTable) {
                this.fields.set(fieldVariants.fullColumnName, colNum);

                if (fieldVariants.fullColumnAliasName !== "") {
                    this.fields.set(fieldVariants.fullColumnAliasName, colNum);
                }
            }
        }
    }
}

const DERIVEDTABLE = "::DERIVEDTABLE::";

/** Perform SQL SELECT operations to retrieve requested data. */
class SelectTables {
    /**
     * @param {Object} ast - Abstract Syntax Tree
     * @param {Map<String,Table>} tableInfo - Map of table info.
     * @param {BindData} bindVariables - List of bind data.
     */
    constructor(ast, tableInfo, bindVariables) {
        /** @property {String} - primary table name. */
        this.primaryTable = ast.FROM.table;

        /** @property {Object} - AST of SELECT fields */
        this.astFields = ast.SELECT;

        /** @property {Map<String,Table>} tableInfo - Map of table info. */
        this.tableInfo = tableInfo;

        /** @property {BindData} - Bind variable data. */
        this.bindVariables = bindVariables;

        /** @property {JoinTables} - Join table object. */
        this.dataJoin = new JoinTables();

        /** @property {TableFields} */
        this.tableFields = new TableFields();

        if (!tableInfo.has(this.primaryTable.toUpperCase()))
            throw new Error(`Invalid table name: ${this.primaryTable}`);

        /** @property {Table} - Primary table info. */
        this.primaryTableInfo = tableInfo.get(this.primaryTable.toUpperCase());

        //  Keep a list of all possible fields from all tables.
        this.tableFields.loadVirtualFields(this.primaryTable, tableInfo);
    }

    /**
     * Update internal FIELDS list to indicate those fields that are in the SELECT fields - that will be returned in data.
     * @param {Object} ast
     * @returns {void} 
     */
    updateSelectedFields(ast) {
        let astFields = ast.SELECT;

        const tableInfo = !this.dataJoin.isDerivedTable() ? this.primaryTableInfo : this.dataJoin.derivedTable.tableInfo;

        //  Expand any 'SELECT *' fields and add the actual field names into 'astFields'.
        astFields = VirtualFields.expandWildcardFields(tableInfo, astFields);

        //  Define the data source of each field in SELECT field list.
        this.tableFields.updateSelectFieldList(astFields, 0, false);

        //  These are fields REFERENCED, but not actually in the SELECT FIELDS.
        //  So columns referenced by GROUP BY, ORDER BY and not in SELECT.
        //  These temp columns need to be removed after processing.
        if (typeof ast["GROUP BY"] !== 'undefined') {
            this.tableFields.updateSelectFieldList(ast["GROUP BY"], this.tableFields.getNextSelectColumnNumber(), true);
        }

        if (typeof ast["ORDER BY"] !== 'undefined') {
            this.tableFields.updateSelectFieldList(ast["ORDER BY"], this.tableFields.getNextSelectColumnNumber(), true);
        }
    }

    /**
     * Process any JOIN condition.
     * @param {Object} ast - Abstract Syntax Tree
     * @returns {void}
     */
    join(ast) {
        if (typeof ast.JOIN !== 'undefined')
            this.dataJoin.load(ast.JOIN, this.tableFields);
    }

    /**
      * Retrieve filtered record ID's.
      * @param {Object} ast - Abstract Syntax Tree
      * @returns {Number[]} - Records ID's that match WHERE condition.
      */
    whereCondition(ast) {
        let sqlData = [];

        let conditions = {};
        if (typeof ast.WHERE !== 'undefined') {
            conditions = ast.WHERE;
        }
        else {
            //  Entire table is selected.  
            conditions = { operator: "=", left: "\"A\"", right: "\"A\"" };
        }

        if (typeof conditions.logic === 'undefined')
            sqlData = this.resolveCondition("OR", [conditions]);
        else
            sqlData = this.resolveCondition(conditions.logic, conditions.terms);

        return sqlData;
    }

    /**
    * Recursively resolve WHERE condition and then apply AND/OR logic to results.
    * @param {String} logic - logic condition (AND/OR) between terms
    * @param {Object} terms - terms of WHERE condition (value compared to value)
    * @returns {Number[]} - record ID's 
    */
    resolveCondition(logic, terms) {
        const recordIDs = [];

        for (const cond of terms) {
            if (typeof cond.logic === 'undefined') {
                recordIDs.push(this.getRecordIDs(cond));
            }
            else {
                recordIDs.push(this.resolveCondition(cond.logic, cond.terms));
            }
        }

        let result = [];
        if (logic === "AND") {
            result = recordIDs.reduce((a, b) => a.filter(c => b.includes(c)));
        }
        if (logic === "OR") {
            //  OR Logic
            let tempArr = [];
            for (const arr of recordIDs) {
                tempArr = tempArr.concat(arr);
            }
            result = Array.from(new Set(tempArr));
        }

        return result;
    }

    /**
    * Find record ID's where condition is TRUE.
    * @param {Object} condition - WHERE test condition
    * @returns {Number[]} - record ID's which are true.
    */
    getRecordIDs(condition) {
        /** @type {Number[]} */
        const recordIDs = [];

        const leftFieldConditions = this.resolveFieldCondition(condition.left);
        const rightFieldConditions = this.resolveFieldCondition(condition.right);

        /** @type {Table} */
        this.masterTable = this.dataJoin.isDerivedTable() ? this.dataJoin.getJoinedTableInfo() : this.primaryTableInfo;
        const calcSqlField = new CalculatedField(this.masterTable, this.primaryTableInfo, this.tableFields);

        for (let masterRecordID = 1; masterRecordID < this.masterTable.tableData.length; masterRecordID++) {
            let leftValue = SelectTables.getConditionValue(leftFieldConditions, calcSqlField, masterRecordID);
            let rightValue = SelectTables.getConditionValue(rightFieldConditions, calcSqlField, masterRecordID);

            if (leftValue instanceof Date || rightValue instanceof Date) {
                leftValue = SelectTables.dateToMs(leftValue);
                rightValue = SelectTables.dateToMs(rightValue);
            }

            if (SelectTables.isConditionTrue(leftValue, condition.operator, rightValue))
                recordIDs.push(masterRecordID);
        }

        return recordIDs;
    }

    /**
     * Evaulate value on left/right side of condition
     * @param {ResolvedFieldCondition} fieldConditions - the value to be found will come from:
     * * constant data
     * * field data
     * * calculated field
     * * sub-query 
     * @param {CalculatedField} calcSqlField - data to resolve the calculated field.
     * @param {Number} masterRecordID - current record in table to grab field data from
     * @returns {any} - resolve value.
     */
    static getConditionValue(fieldConditions, calcSqlField, masterRecordID) {
        let leftValue = fieldConditions.constantData;
        if (fieldConditions.columnNumber >= 0) {
            leftValue = fieldConditions.fieldConditionTableInfo.tableData[masterRecordID][fieldConditions.columnNumber];
        }
        else if (fieldConditions.calculatedField !== "") {
            if (fieldConditions.calculatedField.toUpperCase() === "NULL") {
                leftValue = "NULL";
            }
            else {
                leftValue = calcSqlField.evaluateCalculatedField(fieldConditions.calculatedField, masterRecordID);
            }
        }
        else if (fieldConditions.subQuery !== null) {
            const arrayResult = fieldConditions.subQuery.select(masterRecordID, calcSqlField);
            if (typeof arrayResult !== 'undefined' && arrayResult !== null && arrayResult.length > 0)
                leftValue = arrayResult[0][0];
        }

        return leftValue;
    }

    /**
     * Compare where term values using operator and see if comparision is true.
     * @param {any} leftValue - left value of condition
     * @param {String} operator - operator for comparision
     * @param {any} rightValue  - right value of condition
     * @returns {Boolean} - is comparison true.
     */
    static isConditionTrue(leftValue, operator, rightValue) {
        let keep = false;

        switch (operator.toUpperCase()) {
            case "=":
                keep = leftValue == rightValue;         // skipcq: JS-0050
                break;

            case ">":
                keep = leftValue > rightValue;
                break;

            case "<":
                keep = leftValue < rightValue;
                break;

            case ">=":
                keep = leftValue >= rightValue;
                break;

            case "<=":
                keep = leftValue <= rightValue;
                break;

            case "<>":
                keep = leftValue != rightValue;         // skipcq: JS-0050
                break;

            case "!=":
                keep = leftValue != rightValue;         // skipcq: JS-0050
                break;

            case "LIKE":
                keep = SelectTables.likeCondition(leftValue, rightValue);
                break;

            case "NOT LIKE":
                keep = SelectTables.notLikeCondition(leftValue, rightValue);
                break;

            case "IN":
                keep = SelectTables.inCondition(leftValue, rightValue);
                break;

            case "NOT IN":
                keep = !(SelectTables.inCondition(leftValue, rightValue));
                break;

            case "IS NOT":
                keep = !(SelectTables.isCondition(leftValue, rightValue));
                break;

            case "IS":
                keep = SelectTables.isCondition(leftValue, rightValue);
                break;

            case "EXISTS":
                keep = SelectTables.existsCondition(rightValue);
                break;

            case "NOT EXISTS":
                keep = !(SelectTables.existsCondition(rightValue));
                break;

            default:
                throw new Error(`Invalid Operator: ${operator}`);
        }

        return keep;
    }

    /**
     * Retrieve the data for the record ID's specified for ALL SELECT fields.
     * @param {Number[]} recordIDs - record ID's which are SELECTed.
     * @returns {any[][]} - double array of select data.  No column title is included here.
     */
    getViewData(recordIDs) {
        const virtualData = [];
        const calcSqlField = new CalculatedField(this.masterTable, this.primaryTableInfo, this.tableFields);
        const subQuery = new CorrelatedSubQuery(this.tableInfo, this.tableFields, this.bindVariables);

        for (const masterRecordID of recordIDs) {
            const newRow = [];

            for (const field of this.tableFields.getSelectFields()) {
                if (field.tableInfo !== null)
                    newRow.push(field.getData(masterRecordID));
                else if (field.subQueryAst !== null) {
                    const result = subQuery.select(masterRecordID, calcSqlField, field.subQueryAst);
                    newRow.push(result[0][0]);
                }
                else if (field.calculatedFormula !== "") {
                    const result = calcSqlField.evaluateCalculatedField(field.calculatedFormula, masterRecordID);
                    newRow.push(result);
                }
            }

            virtualData.push(newRow);
        }

        return virtualData;
    }

    /**
     * Returns the entire string in UPPER CASE - except for anything between quotes.
     * @param {String} srcString - source string to convert.
     * @returns {String} - converted string.
     */
    static toUpperCaseExceptQuoted(srcString) {
        let finalString = "";
        let inQuotes = "";

        for (let i = 0; i < srcString.length; i++) {
            let ch = srcString.charAt(i);

            if (inQuotes === "") {
                if (ch === '"' || ch === "'")
                    inQuotes = ch;
                ch = ch.toUpperCase();
            }
            else {
                if (ch === inQuotes)
                    inQuotes = "";
            }

            finalString += ch;
        }

        return finalString;
    }

    /**
     * Parse input string for 'func' and then parse if found.
     * @param {String} functionString - Select field which may contain a function.
     * @param {String} func - Function name to parse for.
     * @returns {String[]} - Parsed function string.
     *   * null if function not found, 
     *   * string array[0] - original string, e.g. **sum(quantity)**
     *   * string array[1] - function parameter, e.g. **quantity**
     */
    static parseForFunctions(functionString, func) {
        const args = [];
        const expMatch = "%1\\s*\\(";

        const matchStr = new RegExp(expMatch.replace("%1", func));
        const startMatchPos = functionString.search(matchStr);
        if (startMatchPos !== -1) {
            const searchStr = functionString.substring(startMatchPos);
            let i = searchStr.indexOf("(");
            const startLeft = i;
            let leftBracket = 1;
            for (i = i + 1; i < searchStr.length; i++) {
                const ch = searchStr.charAt(i);
                if (ch === "(") leftBracket++;
                if (ch === ")") leftBracket--;

                if (leftBracket === 0) {
                    args.push(searchStr.substring(0, i + 1));
                    args.push(searchStr.substring(startLeft + 1, i));
                    return args;
                }
            }
        }

        return null;
    }

    /**
     * Parse the input for a calculated field.
     * String split on comma, EXCEPT if comma is within brackets (i.e. within an inner function)
     * @param {String} paramString - Search and parse this string for parameters.
     * @returns {String[]} - List of function parameters.
     */
    static parseForParams(paramString, startBracket = "(", endBracket = ")") {
        const args = [];
        let bracketCount = 0;
        let start = 0;

        for (let i = 0; i < paramString.length; i++) {
            const ch = paramString.charAt(i);

            if (ch === "," && bracketCount === 0) {
                args.push(paramString.substring(start, i));
                start = i + 1;
            }
            else if (ch === startBracket)
                bracketCount++;
            else if (ch === endBracket)
                bracketCount--;
        }

        const lastStr = paramString.substring(start);
        if (lastStr !== "")
            args.push(lastStr);

        return args;
    }

    /**
     * Compress the table data so there is one record per group (fields in GROUP BY).
     * The other fields MUST be aggregate calculated fields that works on the data in that group.
     * @param {Object} ast - Abstract Syntax Tree
     * @param {any[][]} viewTableData - Table data.
     * @returns {any[][]} - Aggregated table data.
     */
    groupBy(ast, viewTableData) {
        let groupedTableData = viewTableData;

        if (typeof ast['GROUP BY'] !== 'undefined') {
            groupedTableData = this.groupByFields(ast['GROUP BY'], viewTableData);

            if (typeof ast.HAVING !== 'undefined') {
                groupedTableData = this.having(ast.HAVING, groupedTableData);
            }
        }
        else {
            //  If any conglomerate field functions (SUM, COUNT,...)
            //  we summarize all records into ONE.
            if (this.tableFields.getConglomerateFieldCount() > 0) {
                const compressedData = [];
                const conglomerate = new ConglomerateRecord(this.tableFields.getSelectFields());
                compressedData.push(conglomerate.squish(viewTableData));
                groupedTableData = compressedData;
            }
        }

        return groupedTableData;
    }

    /**
    * Group table data by group fields.
    * @param {any[]} astGroupBy - AST group by fields.
    * @param {any[][]} selectedData - table data
    * @returns {any[][]} - compressed table data
    */
    groupByFields(astGroupBy, selectedData) {
        if (selectedData.length === 0)
            return selectedData;

        //  Sort the least important first, and most important last.
        astGroupBy.reverse();

        for (const orderField of astGroupBy) {
            const selectColumn = this.tableFields.getSelectFieldColumn(orderField.name);
            if (selectColumn !== -1) {
                SelectTables.sortByColumnASC(selectedData, selectColumn);
            }
        }

        const groupedData = [];
        let groupRecords = [];
        const conglomerate = new ConglomerateRecord(this.tableFields.getSelectFields());

        let lastKey = this.createGroupByKey(selectedData[0], astGroupBy);
        for (const row of selectedData) {
            const newKey = this.createGroupByKey(row, astGroupBy);
            if (newKey !== lastKey) {
                groupedData.push(conglomerate.squish(groupRecords));

                lastKey = newKey;
                groupRecords = [];
            }
            groupRecords.push(row);
        }

        if (groupRecords.length > 0)
            groupedData.push(conglomerate.squish(groupRecords));

        return groupedData;
    }

    /**
     * Create a composite key that is comprised from all field data in group by clause.
     * @param {any[]} row  - current row of data.
     * @param {any[]} astGroupBy - group by fields
     * @returns {String} - group key
     */
    createGroupByKey(row, astGroupBy) {
        let key = "";

        for (const orderField of astGroupBy) {
            const selectColumn = this.tableFields.getSelectFieldColumn(orderField.name);
            if (selectColumn !== -1)
                key += row[selectColumn].toString();
        }

        return key;
    }

    /**
    * Take the compressed data from GROUP BY and then filter those records using HAVING conditions.
    * @param {Object} astHaving - AST HAVING conditons
    * @param {any[][]} selectedData - compressed table data (from group by)
    * @returns {any[][]} - filtered data using HAVING conditions.
    */
    having(astHaving, selectedData) {
        //  Add in the title row for now
        selectedData.unshift(this.tableFields.getColumnNames());

        //  Create our virtual GROUP table with data already selected.
        const groupTable = new Table(this.primaryTable).loadArrayData(selectedData);

        /** @type {Map<String, Table>} */
        const tableMapping = new Map();
        tableMapping.set(this.primaryTable.toUpperCase(), groupTable);

        //  Set up for our SQL.
        const inSQL = new Sql().setTables(tableMapping);

        //  Fudge the HAVING to look like a SELECT.
        const astSelect = {};
        astSelect.FROM = { table: this.primaryTable, as: '' };
        astSelect.SELECT = [{ name: "*" }];
        astSelect.WHERE = astHaving;

        return inSQL.execute(astSelect);
    }

    /**
     * Take select data and sort by columns specified in ORDER BY clause.
     * @param {Object} ast - Abstract Syntax Tree for SELECT
     * @param {any[][]} selectedData - Table data to sort.  On function return, this array is sorted.
     */
    orderBy(ast, selectedData) {
        if (typeof ast['ORDER BY'] === 'undefined')
            return;

        const astOrderby = ast['ORDER BY']

        //  Sort the least important first, and most important last.
        const reverseOrderBy = astOrderby.reverse();

        for (const orderField of reverseOrderBy) {
            const selectColumn = this.tableFields.getSelectFieldColumn(orderField.name);

            if (selectColumn === -1) {
                throw new Error(`Invalid ORDER BY: ${orderField.name}`);
            }

            if (orderField.order.toUpperCase() === "DESC") {
                SelectTables.sortByColumnDESC(selectedData, selectColumn);
            }
            else {
                SelectTables.sortByColumnASC(selectedData, selectColumn);
            }
        }
    }

    /**
     * Removes temporary fields from return data.  These temporary fields were needed to generate
     * the final table data, but are not included in the SELECT fields for final output.
     * @param {any[][]} viewTableData - table data that may contain temporary columns.
     * @returns {any[][]} - table data with temporary columns removed.
     */
    removeTempColumns(viewTableData) {
        const tempColumns = this.tableFields.getTempSelectedColumnNumbers();

        if (tempColumns.length === 0)
            return viewTableData;

        for (const row of viewTableData) {
            for (const col of tempColumns) {
                row.splice(col, 1);
            }
        }

        return viewTableData;
    }

    /**
     * Sort the table data from lowest to highest using the data in colIndex for sorting.
     * @param {any[][]} tableData - table data to sort.
     * @param {Number} colIndex - column index which indicates which column to use for sorting.
     * @returns {any[][]} - sorted table data.
     */
    static sortByColumnASC(tableData, colIndex) {
        tableData.sort(sortFunction);

        /**
         * 
         * @param {any} a 
         * @param {any} b 
         * @returns {Number}
         */
        function sortFunction(a, b) {
            if (a[colIndex] === b[colIndex]) {
                return 0;
            }
            return (a[colIndex] < b[colIndex]) ? -1 : 1;
        }

        return tableData;
    }

    /**
     * Sort the table data from highest to lowest using the data in colIndex for sorting.
     * @param {any[][]} tableData - table data to sort.
     * @param {Number} colIndex - column index which indicates which column to use for sorting.
     * @returns {any[][]} - sorted table data.
     */
    static sortByColumnDESC(tableData, colIndex) {

        tableData.sort(sortFunction);

        /**
         * 
         * @param {any} a 
         * @param {any} b 
         * @returns {Number}
         */
        function sortFunction(a, b) {
            if (a[colIndex] === b[colIndex]) {
                return 0;
            }
            return (a[colIndex] > b[colIndex]) ? -1 : 1;
        }

        return tableData;
    }

    /**
     * @typedef {Object} ResolvedFieldCondition
     * @property {Table} fieldConditionTableInfo
     * @property {Number} columnNumber - use column data from this column, unless -1.
     * @property {String} constantData - constant data used for column, unless null.
     * @property {String} calculatedField - calculation of data for column, unless empty.
     * @property {CorrelatedSubQuery} subQuery - use this correlated subquery object if not null. 
     * 
     */
    /**
     * Determine what the source of value is for the current field condition.
     * @param {Object} fieldCondition - left or right portion of condition
     * @returns {ResolvedFieldCondition}
     */
    resolveFieldCondition(fieldCondition) {
        /** @type {String} */
        let constantData = null;
        /** @type {Number} */
        let columnNumber = -1;
        /** @type {Table} */
        let fieldConditionTableInfo = null;
        /** @type {String} */
        let calculatedField = "";
        /** @type {CorrelatedSubQuery} */
        let subQuery = null;

        if (typeof fieldCondition.SELECT !== 'undefined') {
            //  Maybe a SELECT within...
            [subQuery, constantData] = this.resolveSubQuery(fieldCondition);
        }
        else if (SelectTables.isStringConstant(fieldCondition))
            //  String constant
            constantData = SelectTables.extractStringConstant(fieldCondition);
        else if (fieldCondition.startsWith('?')) {
            //  Bind variable data.
            constantData = this.resolveBindData(fieldCondition);
        }
        else if (!isNaN(fieldCondition)) {
            //  Literal number.
            constantData = fieldCondition;
        }
        else if (this.tableFields.hasField(fieldCondition)) {
            //  Table field.
            columnNumber = this.tableFields.getFieldColumn(fieldCondition);
            fieldConditionTableInfo = this.tableFields.getTableInfo(fieldCondition);
        }
        else {
            //  Calculated field?
            calculatedField = fieldCondition;
        }

        return { fieldConditionTableInfo, columnNumber, constantData, calculatedField, subQuery };
    }

    /**
     * Handle subquery.  If correlated subquery, return object to handle, otherwise resolve and return constant data.
     * @param {Object} fieldCondition - left or right portion of condition
     * @returns {any[]}
     */
    resolveSubQuery(fieldCondition) {
        /** @type {CorrelatedSubQuery} */
        let subQuery = null;
        /** @type {String} */
        let constantData = null;

        if (SelectTables.isCorrelatedSubQuery(fieldCondition)) {
            subQuery = new CorrelatedSubQuery(this.tableInfo, this.tableFields, this.bindVariables, fieldCondition);
        }
        else {
            const subQueryTableInfo = SelectTables.getSubQueryTableSet(fieldCondition, this.tableInfo);
            const inData = new Sql()
                .setTables(subQueryTableInfo)
                .setBindValues(this.bindVariables)
                .execute(fieldCondition);

            constantData = inData.join(",");
        }

        return [subQuery, constantData];
    }

    /**
     * Get constant bind data
     * @param {Object} fieldCondition - left or right portion of condition
     * @returns {any}
     */
    resolveBindData(fieldCondition) {
        //  Bind variable data.
        const constantData = this.bindVariables.get(fieldCondition);
        if (typeof constantData === 'undefined') {
            if (fieldCondition === '?') {
                throw new Error("Bind variable naming is ?1, ?2... where ?1 is first bind data point in list.")
            }
            else {
                throw new Error(`Bind variable ${fieldCondition} was not found`);
            }
        }

        return constantData;
    }

    static isCorrelatedSubQuery(ast) {
        const tableSet = new Map();
        Sql.extractAstTables(ast, tableSet);

        const tableSetCorrelated = new Map();
        if (typeof ast.WHERE !== 'undefined') {
            Sql.getTableNamesWhereCondition(ast.WHERE, tableSetCorrelated);
        }

        // @ts-ignore
        for (const tableName of tableSetCorrelated.keys()) {
            let isFound = false;
            // @ts-ignore
            for (const outerTable of tableSet.keys()) {
                if (outerTable === tableName || tableSet.get(outerTable) === tableName) {
                    isFound = true;
                    break;
                }
            }
            if (!isFound) {
                return true;
            }
        }

        return false;
    }

    /**
     * Create a set of tables that are used in sub-query.
     * @param {Object} ast - Sub-query AST.
     * @param {Map<String,Table>} tableInfo - Master set of tables used for entire select.
     * @returns {Map<String,Table>} - table set for sub-query.
     */
    static getSubQueryTableSet(ast, tableInfo) {
        const tableSubSet = new Map();
        const selectTables = Sql.getReferencedTableNamesFromAst(ast);

        for (const found of selectTables) {
            if (found[0] !== "" && !tableSubSet.has(found[0])) {
                tableSubSet.set(found[0], tableInfo.get(found[0]));
            }
            if (found[1] !== "" && !tableSubSet.has(found[1])) {
                tableSubSet.set(found[1], tableInfo.get(found[1]));
            }
        }

        return tableSubSet;
    }

    /**
     * Is the string a constant in the SELECT condition.  
     * @param {String} value - condition to test
     * @returns {Boolean} - Is this string a constant.
     */
    static isStringConstant(value) {
        return value.startsWith('"') && value.endsWith('"') || value.startsWith("'") && value.endsWith("'");
    }

    /**
     * Extract the string literal out of condition.  This removes surrounding quotes.
     * @param {String} value - String that encloses literal string data.
     * @returns {String} - String with quotes removed.
     */
    static extractStringConstant(value) {
        if (value.startsWith('"') && value.endsWith('"'))
            return value.replace(/"/g, '');

        if (value.startsWith("'") && value.endsWith("'"))
            return value.replace(/'/g, '');

        return value;
    }

    /**
     * Convert input into milliseconds.
     * @param {any} value - date as as Date or String.
     * @returns {Number} - date as ms.
     */
    static dateToMs(value) {
        let year = 0;
        let month = 0;
        let dayNum = 0;

        if (value instanceof Date) {
            year = value.getFullYear();
            month = value.getMonth();
            dayNum = value.getDate();
        }
        else if (typeof value === "string") {
            const dateParts = value.split("/");
            if (dateParts.length === 3) {
                year = Number(dateParts[2]);
                month = Number(dateParts[0]) - 1;
                dayNum = Number(dateParts[1]);
            }
        }

        const newDate = new Date(Date.UTC(year, month, dayNum, 12, 0, 0, 0));
        return newDate.getTime();
    }

    /**
     * Compare strings in LIKE condition
     * @param {String} leftValue - string for comparison
     * @param {String} rightValue - string with wildcard
     * @returns {Boolean} - Do strings match?
     */
    static likeCondition(leftValue, rightValue) {
        if ((leftValue === null || rightValue === null) && !(leftValue === null && rightValue === null)) {
            return false;
        }

        // @ts-ignore
        const expanded = rightValue.replace(/%/g, ".*").replace(/_/g, ".");

        const result = leftValue.search(expanded);
        return result !== -1;
    }

    static notLikeCondition(leftValue, rightValue) {
        if ((leftValue === null || rightValue === null) && !(leftValue === null && rightValue === null)) {
            return false;
        }

        // @ts-ignore
        const expanded = rightValue.replace(/%/g, ".*").replace(/_/g, ".");

        const result = leftValue.search(expanded);
        return result === -1;
    }

    /**
     * Check if leftValue is contained in list in rightValue
     * @param {any} leftValue - value to find in right value
     * @param {String} rightValue - list of comma separated values
     * @returns {Boolean} - Is contained IN list.
     */
    static inCondition(leftValue, rightValue) {
        let items = [];
        if (typeof rightValue === 'string') {
            items = rightValue.split(",");
        }
        else {
            //  select * from table WHERE IN (select number from table)
            // @ts-ignore
            items = [rightValue.toString()];
        }

        for (let i = 0; i < items.length; i++)
            items[i] = items[i].trimStart().trimEnd();

        let index = items.indexOf(leftValue);
        if (index === -1 && typeof leftValue === 'number') {
            index = items.indexOf(leftValue.toString());
        }

        return index !== -1;
    }

    /**
     * If leftValue is empty (we will consider that as NULL), condition will be true
     * @param {any} leftValue - test this value for NULL
     * @param {any} rightValue - 'NULL' considered as NULL.
     * @returns {Boolean} - Is leftValue NULL (like).
     */
    static isCondition(leftValue, rightValue) {
        return (leftValue === "" && rightValue === "NULL");
    }

    /**
     * Test if input is not empty
     * @param {*} rightValue - value to check if empty
     * @returns - true if NOT empty
     */
    static existsCondition(rightValue) {
        return rightValue !== '';
    }

    /**
     * Return a list of column titles for this table.
     * @param {String} columnTableNameReplacement
     * @returns {String[]} - column titles
     */
    getColumnTitles(columnTableNameReplacement) {
        return this.tableFields.getColumnTitles(columnTableNameReplacement);
    }
}

/** Evaulate calculated fields in SELECT statement.  This is achieved by converting the request 
 * into javascript and then using 'Function' to evaulate it.  
 */
class CalculatedField {
    /**
     * 
     * @param {Table} masterTable - JOINed table (unless not joined, then primary table)
     * @param {Table} primaryTable - First table in SELECT
     * @param {TableFields} tableFields - All fields from all tables
     */
    constructor(masterTable, primaryTable, tableFields) {
        /** @property {Table} */
        this.masterTable = masterTable;
        /** @property {Table} */
        this.primaryTable = primaryTable;
        /** @property {Map<String,String>} - Map key=calculated field in SELECT, value=javascript equivalent code */
        this.sqlServerFunctionCache = new Map();
        /** @property {TableField[]} */
        this.masterFields = tableFields.allFields.filter((vField) => this.masterTable === vField.tableInfo);

        this.mapMasterFields = new Map();
        for (const fld of this.masterFields) {
            this.mapMasterFields.set(fld.fieldName, fld);
        }
    }

    /**
     * Get data from the table for the requested field name and record number
     * @param {String} fldName - Name of field to get data for.
     * @param {Number} masterRecordID - The row number in table to extract data from.
     * @returns {any} - Data from table.  undefined if not found.
     */
    getData(fldName, masterRecordID) {
        const vField = this.mapMasterFields.get(fldName);
        if (typeof vField === 'undefined')
            return vField;

        return vField.getData(masterRecordID);
    }

    /**
     * Evaluate the calculated field for the current table record and return a value.
     * @param {String} calculatedFormula - calculation from SELECT statement
     * @param {Number} masterRecordID - current record ID.
     * @returns {any} - Evaluated data from calculation.
     */
    evaluateCalculatedField(calculatedFormula, masterRecordID) {
        let result = "";

        // e.g.  special case.  count(*)
        if (calculatedFormula === "*") {
            return "*";
        }

        const functionString = this.sqlServerCalcFields(calculatedFormula, masterRecordID);
        try {
            result = new Function(functionString)();
        }
        catch (ex) {
            throw new Error(`Calculated Field Error: ${ex.message}.  ${functionString}`);
        }

        return result;
    }

    /**
     * The program is attempting to build some javascript code which we can then execute to 
     * find the value of the calculated field.  There are two parts.
     * 1)  Build LET statements to assign to all possible field name variants,
     * 2)  Add the 'massaged' calculated field so that it can be run in javascript.
     * @param {String} calculatedFormula - calculation from SELECT statement
     * @param {Number} masterRecordID - current table record ID.
     * @returns {String} - String to be executed.  It is valid javascript lines of code.
     */
    sqlServerCalcFields(calculatedFormula, masterRecordID) {
        //  Working on a calculated field.
        const objectsDeclared = new Map();
        const variablesDeclared = new Map();

        let myVars = "";
        for (/** @type {TableField} */ const vField of this.masterFields) {
            //  Get the DATA from this field.  We then build a series of LET statments
            //  and we assign that data to the field name that might be found in a calculated field.
            let varData = vField.getData(masterRecordID);
            if (varData instanceof Date) {
                varData = `'${varData}'`;
            }
            else if (typeof varData === "string") {
                varData = varData.replace(/'/g, "\\'");
                varData = `'${varData}'`;
            }

            myVars += this.createAssignmentStatments(vField, objectsDeclared, variablesDeclared, varData);
        }

        const functionString = this.sqlServerFunctions(calculatedFormula);

        return `${myVars} return ${functionString}`;
    }

    /**
     * Creates a javascript code block.  For the current field (vField), a variable is assigned the appropriate
     * value from 'varData'.  For example, if the column was 'ID' and the table was 'BOOKS'.
     * ```
     * "let BOOKS = {};BOOKS.ID = '9';"
     * ```
     * If the BOOKS object had already been declared, later variables would just be:
     * ```
     * "BOOKS.NAME = 'To Kill a Blue Jay';"
     * ```
     * @param {TableField} vField - current field that LET statements will be assigning to.
     * @param {Map<String, Boolean>} objectsDeclared - tracks if TABLE name was been encountered yet.
     * @param {Map<String, Boolean>} variablesDeclared - tracks if variables has already been assigned.
     * @param {String} varData - the data from the table that will be assigned to the variable.
     * @returns {String} - the javascript code block.
     */
    createAssignmentStatments(vField, objectsDeclared, variablesDeclared, varData) {
        let myVars = "";

        for (const aliasName of vField.aliasNames) {
            if ((this.primaryTable.tableName !== vField.tableInfo.tableName && aliasName.indexOf(".") === -1))
                continue;

            if (aliasName.indexOf(".") === -1) {
                if (!variablesDeclared.has(aliasName)) {
                    myVars += `let ${aliasName} = ${varData};`;
                    variablesDeclared.set(aliasName, true);
                }
            }
            else {
                const parts = aliasName.split(".");
                if (!objectsDeclared.has(parts[0])) {
                    myVars += `let ${parts[0]} = {};`;
                    objectsDeclared.set(parts[0], true);
                }
                myVars += `${aliasName} = ${varData};`;
            }
        }

        return myVars;
    }

    /**
     * 
     * @param {String} calculatedFormula 
     * @returns {String}
     */
    sqlServerFunctions(calculatedFormula) {
        //  If this calculated field formula has already been put into the required format,
        //  pull this out of our cache rather than redo.
        if (this.sqlServerFunctionCache.has(calculatedFormula))
            return this.sqlServerFunctionCache.get(calculatedFormula);

        const func = new SqlServerFunctions();
        const functionString = func.convertToJs(calculatedFormula, this.masterFields);

        //  No need to recalculate for each row.
        this.sqlServerFunctionCache.set(calculatedFormula, functionString);

        return functionString;
    }
}

/** Correlated Sub-Query requires special lookups for every record in the primary table. */
class CorrelatedSubQuery {
    /**
     * 
     * @param {Map<String, Table>} tableInfo - Map of table info.
     * @param {TableFields} tableFields - Fields from all tables.
     * @param {BindData}  bindData - List of bind data.
     * @param {Object} defaultSubQuery - Select AST
     */
    constructor(tableInfo, tableFields, bindData, defaultSubQuery = null) {
        /** @property {Map<String, Table>} - Map of table info. */
        this.tableInfo = tableInfo;
        /** @property {TableFields} - Fields from all tables.*/
        this.tableFields = tableFields;
        /** @property {BindData} */
        this.bindVariables = bindData;
        /** @property {Object} - AST can be set here and skipped in select() statement. */
        this.defaultSubQuery = defaultSubQuery;
    }

    /**
     * Perform SELECT on sub-query using data from current record in outer table.
     * @param {Number} masterRecordID - Current record number in outer table.
     * @param {CalculatedField} calcSqlField - Calculated field object.
     * @param {Object} ast - Sub-query AST.
     * @returns {any[][]} - double array of selected table data.
     */
    select(masterRecordID, calcSqlField, ast = this.defaultSubQuery) {
        const innerTableInfo = this.tableInfo.get(ast.FROM.table.toUpperCase());
        if (typeof innerTableInfo === 'undefined')
            throw new Error(`No table data found: ${ast.FROM.table}`);

        //  Add BIND variable for all matching fields in WHERE.
        const tempAst = JSON.parse(JSON.stringify(ast));
        const tempBindVariables = new BindData();
        tempBindVariables.addList(this.bindVariables.getBindDataList());

        this.replaceOuterFieldValueInCorrelatedWhere(calcSqlField, masterRecordID, tempAst, tempBindVariables);

        const inData = new Sql()
            .setTables(this.tableInfo)
            .setBindValues(tempBindVariables)
            .execute(tempAst);

        return inData;
    }

    /**
     * If we find the field name in the AST, just replace with '?' and add to bind data variable list.
     * @param {CalculatedField} calcSqlField - List of fields in outer query.  If any are found in subquery, the value of that field for the current record is inserted into subquery before it is executed.
     * @param {Number} masterRecordID - current record number in outer query.
     * @param {Object} tempAst - AST for subquery.  Any field names found from outer query will be replaced with bind place holder '?'.
     * @param {BindData} bindData
     */
    replaceOuterFieldValueInCorrelatedWhere(calcSqlField, masterRecordID, tempAst, bindData) {
        const where = tempAst.WHERE;

        if (typeof where === 'undefined')
            return;

        if (typeof where.logic === 'undefined')
            this.traverseWhere(calcSqlField, [where], masterRecordID, bindData);
        else
            this.traverseWhere(calcSqlField, where.terms, masterRecordID, bindData);
    }

    /**
     * Search the WHERE portion of the subquery to find all references to the table in the outer query.
     * @param {CalculatedField} calcSqlField - List of fields in outer query.
     * @param {Object} terms - terms of WHERE.  It is modified with bind variable placeholders when outer table fields are located.
     * @param {Number} masterRecordID
     * @param {BindData} bindData
     */
    traverseWhere(calcSqlField, terms, masterRecordID, bindData) {

        for (const cond of terms) {
            if (typeof cond.logic === 'undefined') {
                let result = calcSqlField.masterFields.find(item => item.fieldName === cond.left.toUpperCase());
                if (typeof result !== 'undefined') {
                    cond.left = bindData.add(calcSqlField.getData(cond.left.toUpperCase(), masterRecordID));
                }
                result = calcSqlField.masterFields.find(item => item.fieldName === cond.right.toUpperCase());
                if (typeof result !== 'undefined') {
                    cond.right = bindData.add(calcSqlField.getData(cond.right.toUpperCase(), masterRecordID));
                }
            }
            else {
                this.traverseWhere(calcSqlField, [cond.terms], masterRecordID, bindData);
            }
        }
    }
}

/** Tracks all fields in a table (including derived tables when there is a JOIN). */
class VirtualFields {
    constructor() {
        /** @property {Map<String, VirtualField>} - Map to field for fast access. Field name is key. */
        this.virtualFieldMap = new Map();
        /** @property {VirtualField[]} - List of all fields for table. */
        this.virtualFieldList = [];
    }

    /**
     * Adds info for one field into master list of fields for table.
     * @param {VirtualField} field - Information for one field in the table.
     */
    add(field, checkForDuplicates = false) {
        if (checkForDuplicates && this.virtualFieldMap.has(field.fieldName)) {
            throw new Error(`Duplicate field name: ${field.fieldName}`);
        }
        this.virtualFieldMap.set(field.fieldName, field);
        this.virtualFieldList.push(field);
    }

    /**
     * Returns a list of all fields in table.
     * @returns {VirtualField[]}
     */
    getAllVirtualFields() {
        return this.virtualFieldList;
    }

    /**
     * When the wildcard '*' is found in the SELECT, it will add all fields in table to the AST used in the SELECT.
     * @param {Table} masterTableInfo - The wildcard '*' (if found) will add fields from THIS table to the AST.
     * @param {any[]} astFields - existing SELECT fields list.
     * @returns {any[]} - original AST field list PLUS expanded list of fields if '*' was encountered.
     */
    static expandWildcardFields(masterTableInfo, astFields) {
        for (let i = 0; i < astFields.length; i++) {
            if (astFields[i].name === "*") {
                //  Replace wildcard will actual field names from master table.
                const masterTableFields = [];
                const allExpandedFields = masterTableInfo.getAllExtendedNotationFieldNames();

                for (const virtualField of allExpandedFields) {
                    const selField = { name: virtualField };
                    masterTableFields.push(selField);
                }

                astFields.splice(i, 1, ...masterTableFields);
                break;
            }
        }

        return astFields;
    }
}

/**  Defines all possible table fields including '*' and long/short form (i.e. table.column). */
class VirtualField {                        //  skipcq: JS-0128
    /**
     * 
     * @param {String} fieldName - field name
     * @param {Table} tableInfo - table this field belongs to.
     * @param {Number} tableColumn - column number of this field.
     */
    constructor(fieldName, tableInfo, tableColumn) {
        /** @property {String} - field name */
        this.fieldName = fieldName;
        /** @property {Table} - table this field belongs to. */
        this.tableInfo = tableInfo;
        /** @property {Number} - column number of this field. */
        this.tableColumn = tableColumn;
    }
}

/** Handle the various JOIN table types. */
class JoinTables {
    /**
     * Join the tables and create a derived table with the combined data from all.
     * @param {any[]} astJoin - AST list of tables to join.
     * @param {TableFields} tableFields
     */
    load(astJoin, tableFields) {
        /** @property {DerivedTable} - result table after tables are joined */
        this.derivedTable = new DerivedTable();
        this.tableFields = tableFields;

        for (const joinTable of astJoin) {
            this.joinNextTable(joinTable);
        }
    }

    /**
     * Updates derived table with join to new table.
     * @param {Object} astJoin 
     */
    joinNextTable(astJoin) {
        this.leftRightFieldInfo = null;
        const recIds = this.joinCondition(astJoin);

        this.derivedTable = JoinTables.joinTables(this.leftRightFieldInfo, astJoin, recIds);

        //  Field locations have changed to the derived table, so update our
        //  virtual field list with proper settings.
        this.tableFields.updateDerivedTableVirtualFields(this.derivedTable);
    }

    /**
     * 
     * @param {Object} conditions 
     * @returns {Array}
     */
    joinCondition(conditions) {
        let recIds = [];
        const rightTable = conditions.table;
        const joinType = conditions.type;

        if (typeof conditions.cond.logic === 'undefined')
            recIds = this.resolveCondition("OR", [conditions], joinType, rightTable);
        else
            recIds = this.resolveCondition(conditions.cond.logic, conditions.cond.terms, joinType, rightTable);

        return recIds;
    }

    /**
     * 
     * @param {String} logic - AND, OR 
     * @param {Object} astConditions 
     * @param {String} joinType - inner, full, left, right
     * @param {*} rightTable - join table.
     * @returns {Array}
     */
    resolveCondition(logic, astConditions, joinType, rightTable) {
        let leftIds = [];
        let rightIds = [];
        let resultsLeft = [];
        let resultsRight = [];

        for (const cond of astConditions) {
            if (typeof cond.logic === 'undefined') {
                [leftIds, rightIds] = this.getRecordIDs(cond, joinType, rightTable);
                resultsLeft.push(leftIds);
                resultsRight.push(rightIds);
            }
            else {
                [leftIds, rightIds] = this.resolveCondition(cond.logic, cond.terms, joinType, rightTable);
                resultsLeft.push(leftIds);
                resultsRight.push(rightIds);
            }
        }

        if (logic === "AND") {
            resultsLeft = JoinTables.andJoinIds(resultsLeft);
            resultsRight = JoinTables.andJoinIds(resultsRight);
        }
        if (logic === "OR") {
            resultsLeft = JoinTables.orJoinIds(resultsLeft);
            resultsRight = JoinTables.orJoinIds(resultsRight);
        }

        return [resultsLeft, resultsRight];
    }

    /**
     * AND logic applied to the record ID's
     * @param {Array} recIds 
     * @returns {Array}
     */
    static andJoinIds(recIds) {
        const result = [];

        for (let i = 0; i < recIds[0].length; i++) {
            const temp = [];

            for (const rec of recIds) {
                temp.push(typeof rec[i] === 'undefined' ? [] : rec[i]);
            }
            const row = temp.reduce((a, b) => a.filter(c => b.includes(c)));

            if (row.length > 0) {
                result[i] = row;
            }
        }

        return result;
    }

    /**
     * OR logic applied to the record ID's
     * @param {Array} recIds 
     * @returns {Array}
     */
    static orJoinIds(recIds) {
        const result = [];

        for (let i = 0; i < recIds[0].length; i++) {
            let temp = [];

            for (const rec of recIds) {
                temp = temp.concat(rec[i]);
            }

            if (typeof temp[0] !== 'undefined') {
                result[i] = Array.from(new Set(temp));
            }
        }

        return result;
    }

    /**
     * 
     * @param {Object} conditionAst 
     * @param {String} joinType - left, right, inner, full
     * @param {String} rightTable 
     * @returns {Array}
     */
    getRecordIDs(conditionAst, joinType, rightTable) {
        this.leftRightFieldInfo = JoinTables.getLeftRightFieldInfo(conditionAst, this.tableFields, rightTable);
        const recIds = JoinTables.getMatchedRecordIds(joinType, this.leftRightFieldInfo);

        return recIds;
    }

    /**
     * 
     * @param {Object} astJoin 
     * @param {TableFields} tableFields 
     * @returns {TableField[]}
     */
    static getLeftRightFieldInfo(astJoin, tableFields, joinedTable) {
        /** @type {TableField} */
        let leftFieldInfo = null;
        /** @type {TableField} */
        let rightFieldInfo = null;

        const left = typeof astJoin.cond === 'undefined' ? astJoin.left : astJoin.cond.left;
        const right = typeof astJoin.cond === 'undefined' ? astJoin.right : astJoin.cond.right;

        leftFieldInfo = tableFields.getFieldInfo(left);
        rightFieldInfo = tableFields.getFieldInfo(right);
        //  joinTable.table is the RIGHT table, so switch if equal to condition left.
        if (joinedTable === leftFieldInfo.originalTable) {
            leftFieldInfo = tableFields.getFieldInfo(right);
            rightFieldInfo = tableFields.getFieldInfo(left);
        }

        return [leftFieldInfo, rightFieldInfo];
    }

    /**
     * 
     * @param {String} type 
     * @param {TableField[]} leftRightFieldInfo 
     * @returns {Array}
     */
    static getMatchedRecordIds(type, leftRightFieldInfo) {
        /** @type {Number[][]} */
        let matchedRecordIDs = [];
        let rightJoinRecordIDs = [];
        /** @type {TableField} */
        let leftFieldInfo = null;
        /** @type {TableField} */
        let rightFieldInfo = null;

        [leftFieldInfo, rightFieldInfo] = leftRightFieldInfo;

        switch (type) {
            case "left":
                matchedRecordIDs = JoinTables.leftRightJoin(leftFieldInfo, rightFieldInfo, type);
                break;
            case "inner":
                matchedRecordIDs = JoinTables.leftRightJoin(leftFieldInfo, rightFieldInfo, type);
                break;
            case "right":
                matchedRecordIDs = JoinTables.leftRightJoin(rightFieldInfo, leftFieldInfo, type);
                break;
            case "full":
                matchedRecordIDs = JoinTables.leftRightJoin(leftFieldInfo, rightFieldInfo, type);
                rightJoinRecordIDs = JoinTables.leftRightJoin(rightFieldInfo, leftFieldInfo, "outer");
                break;
            default:
                throw new Error(`Invalid join type: ${type}`);
        }

        return [matchedRecordIDs, rightJoinRecordIDs];
    }

    /**
     * Does this object contain a derived (joined) table.
     * @returns {Boolean}
     */
    isDerivedTable() {
        if (typeof this.derivedTable === 'undefined') {
            return false;
        }

        return this.derivedTable.isDerivedTable();
    }

    /**
     * Get derived table after tables are joined.
     * @returns {Table}
     */
    getJoinedTableInfo() {
        return this.derivedTable.getTableData();
    }

    /**
    * Join two tables and create a derived table that contains all data from both tables.
    * @param {TableField[]} leftRightFieldInfo - left table field of join
    * @param {Object} joinTable - AST that contains join type.
    * @param {Array} recIds
    * @returns {DerivedTable} - new derived table after join of left and right tables.
    */
    static joinTables(leftRightFieldInfo, joinTable, recIds) {
        let derivedTable = null;
        let rightDerivedTable = null;

        const [leftFieldInfo, rightFieldInfo] = leftRightFieldInfo;
        const [matchedRecordIDs, rightJoinRecordIDs] = recIds;

        switch (joinTable.type) {
            case "left":
                derivedTable = new DerivedTable()
                    .setLeftField(leftFieldInfo)
                    .setRightField(rightFieldInfo)
                    .setLeftRecords(matchedRecordIDs)
                    .setIsOuterJoin(true)
                    .createTable();
                break;

            case "inner":
                derivedTable = new DerivedTable()
                    .setLeftField(leftFieldInfo)
                    .setRightField(rightFieldInfo)
                    .setLeftRecords(matchedRecordIDs)
                    .setIsOuterJoin(false)
                    .createTable();
                break;

            case "right":
                derivedTable = new DerivedTable()
                    .setLeftField(rightFieldInfo)
                    .setRightField(leftFieldInfo)
                    .setLeftRecords(matchedRecordIDs)
                    .setIsOuterJoin(true)
                    .createTable();

                break;

            case "full":
                derivedTable = new DerivedTable()
                    .setLeftField(leftFieldInfo)
                    .setRightField(rightFieldInfo)
                    .setLeftRecords(matchedRecordIDs)
                    .setIsOuterJoin(true)
                    .createTable();

                rightDerivedTable = new DerivedTable()
                    .setLeftField(rightFieldInfo)
                    .setRightField(leftFieldInfo)
                    .setLeftRecords(rightJoinRecordIDs)
                    .setIsOuterJoin(true)
                    .createTable();

                derivedTable.tableInfo.concat(rightDerivedTable.tableInfo);         // skipcq: JS-D008

                break;

            default:
                throw new Error(`Internal error.  No support for join type: ${joinTable.type}`);
        }
        return derivedTable;
    }

    /**
     * Returns array of each matching record ID from right table for every record in left table.
     * If the right table entry could NOT be found, -1 is set for that record index.
     * @param {TableField} leftField - left table field
     * @param {TableField} rightField - right table field
     * @param {String} type - either 'inner' or 'outer'.
     * @returns {Number[][]} - first index is record ID of left table, second index is a list of the matching record ID's in right table.
     */
    static leftRightJoin(leftField, rightField, type) {
        const leftRecordsIDs = [];

        //  First record is the column title.
        leftRecordsIDs.push([0]);

        /** @type {any[][]} */
        const leftTableData = leftField.tableInfo.tableData;
        const leftTableCol = leftField.tableColumn;

        rightField.tableInfo.addIndex(rightField.fieldName);
        const searchFieldCol = rightField.tableInfo.getFieldColumn(rightField.fieldName);
        const searchName = rightField.fieldName.trim().toUpperCase();

        for (let leftTableRecordNum = 1; leftTableRecordNum < leftTableData.length; leftTableRecordNum++) {
            let keyMasterJoinField = leftTableData[leftTableRecordNum][leftTableCol];
            if (keyMasterJoinField !== null) {
                keyMasterJoinField = keyMasterJoinField.toString();
            }
            const joinRows =  searchFieldCol === -1 ? [] : rightField.tableInfo.search(searchName, keyMasterJoinField);

            //  For the current LEFT TABLE record, record the linking RIGHT TABLE records.
            if (joinRows.length === 0) {
                if (type === "inner")
                    continue;

                leftRecordsIDs[leftTableRecordNum] = [-1];
            }
            else {
                //  Excludes all match recordgs (is outer the right word for this?)
                if (type === "outer")
                    continue;

                leftRecordsIDs[leftTableRecordNum] = joinRows;
            }
        }

        return leftRecordsIDs;
    }
}

/**  The JOIN creates a new logical table. */
class DerivedTable {
    constructor() {
        /** @property {Table} */
        this.tableInfo = null;
        /** @property  {TableField} */
        this.leftField = null;
        /** @property  {TableField} */
        this.rightField = null;
        /** @property  {Number[][]} */
        this.leftRecords = null;
        /** @property  {Boolean} */
        this.isOuterJoin = null;
    }

    /**
     * Left side of join condition.
     * @param {TableField} leftField 
     * @returns {DerivedTable}
     */
    setLeftField(leftField) {
        this.leftField = leftField;
        return this;
    }

    /**
     * Right side of join condition
     * @param {TableField} rightField 
     * @returns {DerivedTable}
     */
    setRightField(rightField) {
        this.rightField = rightField;
        return this;
    }

    /**
     * 
     * @param {Number[][]} leftRecords - first index is record ID of left table, second index is a list of the matching record ID's in right table.
     * @returns {DerivedTable} 
     */
    setLeftRecords(leftRecords) {
        this.leftRecords = leftRecords;
        return this;
    }

    /**
     * Indicate if outer or inner join.
     * @param {Boolean} isOuterJoin - true for outer, false for inner
     * @returns {DerivedTable}
     */
    setIsOuterJoin(isOuterJoin) {
        this.isOuterJoin = isOuterJoin;
        return this;
    }

    /**
     * Create derived table from the two tables that are joined.
     * @returns {DerivedTable}
     */
    createTable() {
        const columnCount = this.rightField.tableInfo.getColumnCount();
        const emptyRightRow = Array(columnCount).fill(null);

        const joinedData = [DerivedTable.getCombinedColumnTitles(this.leftField, this.rightField)];

        for (let i = 1; i < this.leftField.tableInfo.tableData.length; i++) {
            if (typeof this.leftRecords[i] !== "undefined") {
                if (typeof this.rightField.tableInfo.tableData[this.leftRecords[i][0]] === "undefined")
                    joinedData.push(this.leftField.tableInfo.tableData[i].concat(emptyRightRow));
                else {
                    const maxJoin = this.leftRecords[i].length;
                    for (let j = 0; j < maxJoin; j++) {
                        joinedData.push(this.leftField.tableInfo.tableData[i].concat(this.rightField.tableInfo.tableData[this.leftRecords[i][j]]));
                    }
                }
            }
        }
        /** @type {Table} */
        this.tableInfo = new Table(DERIVEDTABLE).loadArrayData(joinedData);

        return this;
    }

    /**
    * Is this a derived table - one that has been joined.
    * @returns {Boolean}
    */
    isDerivedTable() {
        return this.tableInfo !== null;
    }

    /**
     * Get derived table info.
     * @returns {Table}
     */
    getTableData() {
        return this.tableInfo;
    }

    /**
     * Create title row from LEFT and RIGHT table.
     * @param {TableField} leftField 
     * @param {TableField} rightField 
     * @returns {String[]}
     */
    static getCombinedColumnTitles(leftField, rightField) {
        const titleRow = leftField.tableInfo.getAllExtendedNotationFieldNames();
        const rightFieldNames = rightField.tableInfo.getAllExtendedNotationFieldNames();
        return titleRow.concat(rightFieldNames);
    }
}

/** Convert SQL CALCULATED fields into javascript code that can be evaulated and converted to data. */
class SqlServerFunctions {
    /**
     * Convert SQL formula to javascript code.
     * @param {String} calculatedFormula - contains SQL formula and parameter(s)
     * @param {TableField[]} masterFields - table fields
     * @returns {String} - javascript code
     */
    convertToJs(calculatedFormula, masterFields) {
        const sqlFunctions = ["ABS", "CASE", "CEILING", "CHARINDEX", "COALESCE", "CONCAT", "CONCAT_WS", "CONVERT", "DAY", "FLOOR", "IF", "LEFT", "LEN", "LENGTH", "LOG", "LOG10", "LOWER",
            "LTRIM", "MONTH", "NOW", "POWER", "RAND", "REPLICATE", "REVERSE", "RIGHT", "ROUND", "RTRIM",
            "SPACE", "STUFF", "SUBSTR", "SUBSTRING", "SQRT", "TRIM", "UPPER", "YEAR"];
        /** @property {String} - regex to find components of CASE statement. */
        this.matchCaseWhenThenStr = /WHEN(.*?)THEN(.*?)(?=WHEN|ELSE|$)|ELSE(.*?)(?=$)/;
        /** @property {String} - Original CASE statement. */
        this.originalCaseStatement = "";
        /** @property {String} - Existing state of function string when CASE encountered. */
        this.originalFunctionString = "";
        /** @property {Boolean} - when working on each WHEN/THEN in CASE, is this the first one encountered. */
        this.firstCase = true;

        let functionString = SelectTables.toUpperCaseExceptQuoted(calculatedFormula);

        for (const func of sqlFunctions) {
            let args = SelectTables.parseForFunctions(functionString, func);

            [args, functionString] = this.caseStart(func, args, functionString);

            while (args !== null && args.length > 0) {
                // Split on COMMA, except within brackets.
                const parms = typeof args[1] === 'undefined' ? [] : SelectTables.parseForParams(args[1]);

                let replacement = "";
                switch (func) {
                    case "ABS":
                        replacement = `Math.abs(${parms[0]})`;
                        break;
                    case "CASE":
                        replacement = this.caseWhen(args);
                        break;
                    case "CEILING":
                        replacement = `Math.ceil(${parms[0]})`;
                        break;
                    case "CHARINDEX":
                        replacement = SqlServerFunctions.charIndex(parms);
                        break;
                    case "COALESCE":
                        replacement = SqlServerFunctions.coalesce(parms);
                        break;
                    case "CONCAT":
                        replacement = SqlServerFunctions.concat(parms, masterFields);
                        break;
                    case "CONCAT_WS":
                        replacement = SqlServerFunctions.concat_ws(parms, masterFields);
                        break;
                    case "CONVERT":
                        replacement = SqlServerFunctions.convert(parms);
                        break;
                    case "DAY":
                        replacement = `new Date(${parms[0]}).getDate()`;
                        break;
                    case "FLOOR":
                        replacement = `Math.floor(${parms[0]})`;
                        break;
                    case "IF":
                        {
                            const ifCond = SqlParse.sqlCondition2JsCondition(parms[0]);
                            replacement = `${ifCond} ? ${parms[1]} : ${parms[2]};`;
                            break;
                        }
                    case "LEFT":
                        replacement = `${parms[0]}.substring(0,${parms[1]})`;
                        break;
                    case "LEN":
                    case "LENGTH":
                        replacement = `${parms[0]}.length`;
                        break;
                    case "LOG":
                        replacement = `Math.log2(${parms[0]})`;
                        break;
                    case "LOG10":
                        replacement = `Math.log10(${parms[0]})`;
                        break;
                    case "LOWER":
                        replacement = `${parms[0]}.toLowerCase()`;
                        break;
                    case "LTRIM":
                        replacement = `${parms[0]}.trimStart()`;
                        break;
                    case "MONTH":
                        replacement = `new Date(${parms[0]}).getMonth() + 1`;
                        break;
                    case "NOW":
                        replacement = "new Date().toLocaleString()";
                        break;
                    case "POWER":
                        replacement = `Math.pow(${parms[0]},${parms[1]})`;
                        break;
                    case "RAND":
                        replacement = "Math.random()";
                        break;
                    case "REPLICATE":
                        replacement = `${parms[0]}.toString().repeat(${parms[1]})`;
                        break;
                    case "REVERSE":
                        replacement = `${parms[0]}.toString().split("").reverse().join("")`;
                        break;
                    case "RIGHT":
                        replacement = `${parms[0]}.toString().slice(${parms[0]}.length - ${parms[1]})`;
                        break;
                    case "ROUND":
                        replacement = `Math.round(${parms[0]})`;
                        break;
                    case "RTRIM":
                        replacement = `${parms[0]}.toString().trimEnd()`;
                        break;
                    case "SPACE":
                        replacement = `' '.repeat(${parms[0]})`;
                        break;
                    case "STUFF":
                        replacement = `${parms[0]}.toString().substring(0,${parms[1]}-1) + ${parms[3]} + ${parms[0]}.toString().substring(${parms[1]} + ${parms[2]} - 1)`;
                        break;
                    case "SUBSTR":
                    case "SUBSTRING":
                        replacement = `${parms[0]}.toString().substring(${parms[1]} - 1, ${parms[1]} + ${parms[2]} - 1)`;
                        break;
                    case "SQRT":
                        replacement = `Math.sqrt(${parms[0]})`;
                        break;
                    case "TRIM":
                        replacement = `${parms[0]}.toString().trim()`;
                        break;
                    case "UPPER":
                        replacement = `${parms[0]}.toString().toUpperCase()`;
                        break;
                    case "YEAR":
                        replacement = `new Date(${parms[0]}).getFullYear()`;
                        break;
                    default:
                        throw new Error(`Internal Error. Function is missing. ${func}`);
                }

                functionString = functionString.replace(args[0], replacement);

                args = this.parseFunctionArgs(func, functionString);
            }

            functionString = this.caseEnd(func, functionString);
        }

        return functionString;
    }

    /**
     * Search for SELECT function arguments for specified 'func' only.  Special case for 'CASE'.  It breaks down one WHEN condition at a time.
     * @param {String} func - an SQL function name.
     * @param {String} functionString - SELECT SQL string to search
     * @returns {String[]}
     */
    parseFunctionArgs(func, functionString) {
        let args = [];

        if (func === "CASE")
            args = functionString.match(this.matchCaseWhenThenStr);
        else
            args = SelectTables.parseForFunctions(functionString, func);

        return args;
    }

    /**
     * Find the position of a substring within a field - in javascript code.
     * @param {any[]} parms - 
     * * parms[0] - string to search for
     * * parms[1] - field name
     * * parms[2] - start to search from this position (starts at 1)
     * @returns {String} - javascript code to find substring position.
     */
    static charIndex(parms) {
        let replacement = "";

        if (typeof parms[2] === 'undefined')
            replacement = `${parms[1]}.toString().indexOf(${parms[0]}) + 1`;
        else
            replacement = `${parms[1]}.toString().indexOf(${parms[0]},${parms[2]} -1) + 1`;

        return replacement;
    }

    /**
     * Returns first non-empty value in a list, in javascript code.
     * @param {any[]} parms - coalesce parameters - no set limit for number of inputs.
     * @returns {String} - javascript to solve
     */
    static coalesce(parms) {
        let replacement = "";
        for (const parm of parms) {
            replacement += `${parm} !== '' ? ${parm} : `;
        }

        replacement += `''`;

        return replacement;
    }

    /**
     * 
     * @param {any[]} parms 
     * @param {TableField[]} masterFields 
     * @returns {String}
     */
    static concat(parms, masterFields) {
        parms.unshift("''");
        return SqlServerFunctions.concat_ws(parms, masterFields);
    }

    /**
     * Concatenate all data and use separator between concatenated fields.
     * @param {any[]} parms - 
     * * parm[0] - separator string
     * * parms... - data to concatenate.
     * @param {TableField[]} masterFields - fields in table.
     * @returns {String} - javascript to concatenate all data.
     */
    static concat_ws(parms, masterFields) {
        if (parms.length === 0) {
            return "";
        }

        let replacement = "";
        const separator = parms[0];
        let concatFields = [];

        for (let i = 1; i < parms.length; i++) {
            if (parms[i].trim() === "*") {
                const allTableFields = TableField.getAllExtendedAliasNames(masterFields);
                concatFields = concatFields.concat(allTableFields);
            }
            else {
                concatFields.push(parms[i]);
            }
        }

        for (const field of concatFields) {
            if (replacement !== "") {
                replacement += ` + ${separator} + `;
            }

            replacement += `${field}`;
        }

        return replacement;
    }

    /**
     * Convert data to another type.
     * @param {any[]} parms - 
     * * parm[0] - value to convert
     * * parms[1] -  data type.
     * @returns {String} - javascript to convert data to specified type.
     */
    static convert(parms) {
        let replacement = "";

        const dataType = parms[1].toUpperCase().trim();
        switch (dataType) {
            case "SIGNED":
                replacement = `isNaN(parseInt(${parms[0]}, 10))?0:parseInt(${parms[0]}, 10)`;
                break;
            case "DECIMAL":
                replacement = `isNaN(parseFloat(${parms[0]}))?0:parseFloat(${parms[0]})`;
                break;
            case "CHAR":
                replacement = `${parms[0]}.toString()`;
                break;
            default:
                throw new Error(`Unrecognized data type ${dataType} in CONVERT`);
        }

        return replacement;
    }

    /**
     * When examining the SQL Select CASE, parse for next WHEN,END condition.
     * @param {String} func - current function worked on.  If <> 'CASE', ignore.
     * @param {any[]} args - default return value. 
     * @param {String} functionString 
     * @returns {any[]}
     */
    caseStart(func, args, functionString) {
        let caseArguments = args;
        let caseString = functionString;

        if (func === "CASE") {
            caseArguments = functionString.match(/CASE(.*?)END/i);

            if (caseArguments !== null && caseArguments.length > 1) {
                this.firstCase = true;
                this.originalFunctionString = functionString;
                this.originalCaseStatement = caseArguments[0];
                caseString = caseArguments[1];

                caseArguments = caseArguments[1].match(this.matchCaseWhenThenStr);
            }
        }

        return [caseArguments, caseString];
    }

    /**
     * Convert SQL CASE to javascript executeable code to solve case options.
     * @param {any[]} args - current CASE WHEN strings.
     * * args[0] - entire WHEN ... THEN ...
     * * args[1] - parsed string after WHEN, before THEN
     * * args[2] - parse string after THEN
     * @returns {String} - js code to handle this WHEN case.
     */
    caseWhen(args) {
        let replacement = "";

        if (args.length > 2) {
            if (typeof args[1] === 'undefined' && typeof args[2] === 'undefined') {
                replacement = `else return ${args[3]};`;
            }
            else {
                if (this.firstCase) {
                    replacement = "(() => {if (";
                    this.firstCase = false;
                }
                else
                    replacement = "else if (";
                replacement += `${SqlParse.sqlCondition2JsCondition(args[1])}) return ${args[2]} ;`;
            }
        }

        return replacement;
    }

    /**
     * Finish up the javascript code to handle the select CASE.
     * @param {String} func - current function being processed.  If <> 'CASE', ignore.
     * @param {String} funcString - current SQL/javascript string in the process of being converted to js.
     * @returns {String} - updated js code
     */
    caseEnd(func, funcString) {
        let functionString = funcString;

        if (func === "CASE" && this.originalFunctionString !== "") {
            functionString += "})();";      //  end of lambda.
            functionString = this.originalFunctionString.replace(this.originalCaseStatement, functionString);
        }

        return functionString;
    }
}

/** Used to create a single row from multiple rows for GROUP BY expressions. */
class ConglomerateRecord {
    /**
     * 
     * @param {TableField[]} virtualFields 
     */
    constructor(virtualFields) {
        /** @property {TableField[]} */
        this.selectVirtualFields = virtualFields;
    }

    /**
     * Compress group records to a single row by applying appropriate aggregate functions.
     * @param {any[][]} groupRecords - a group of table data records to compress.
     * @returns {any[]} - compressed record.
     * * If column is not an aggregate function, value from first row of group records is selected. (should all be the same)
     * * If column has aggregate function, that function is applied to all rows from group records.
     */
    squish(groupRecords) {
        const row = [];
        if (groupRecords.length === 0)
            return row;

        let i = 0;
        for (/** @type {TableField} */ const field of this.selectVirtualFields) {
            if (field.aggregateFunction === "")
                row.push(groupRecords[0][i]);
            else {
                row.push(ConglomerateRecord.aggregateColumn(field, groupRecords, i));
            }
            i++;
        }
        return row;
    }

    /**
     * Apply aggregate function to all rows on specified column and return result.
     * @param {TableField} field - field with aggregate function
     * @param {any[]} groupRecords - group of records we apply function to.
     * @param {Number} columnIndex - the column index where data is read from and function is applied on.
     * @returns {Number} - value of aggregate function for all group rows.
     */
    static aggregateColumn(field, groupRecords, columnIndex) {
        let groupValue = 0;
        let avgCounter = 0;
        let first = true;
        const distinctSet = new Set();

        for (const groupRow of groupRecords) {
            if (groupRow[columnIndex] === 'null')
                continue;

            let numericData = 0;
            if (groupRow[columnIndex] instanceof Date) {
                numericData = groupRow[columnIndex];
            }
            else {
                numericData = Number(groupRow[columnIndex]);
                numericData = (isNaN(numericData)) ? 0 : numericData;
            }

            switch (field.aggregateFunction) {
                case "SUM":
                    groupValue += numericData;
                    break;
                case "COUNT":
                    groupValue++;
                    if (field.distinctSetting === "DISTINCT") {
                        distinctSet.add(groupRow[columnIndex]);
                        groupValue = distinctSet.size;
                    }
                    break;
                case "MIN":
                    groupValue = ConglomerateRecord.minCase(first, groupValue, numericData);
                    break;
                case "MAX":
                    groupValue = ConglomerateRecord.maxCase(first, groupValue, numericData);
                    break;
                case "AVG":
                    avgCounter++;
                    groupValue += numericData;
                    break;
                default:
                    throw new Error(`Invalid aggregate function: ${field.aggregateFunction}`);
            }
            first = false;
        }

        if (field.aggregateFunction === "AVG")
            groupValue = groupValue / avgCounter;

        return groupValue;
    }

    /**
     * Find minimum value from group records.
     * @param {Boolean} first - true if first record in set.
     * @param {Number} value - cumulative data from all previous group records
     * @param {Number} data - data from current group record
     * @returns {Number} - minimum value from set.
     */
    static minCase(first, value, data) {
        let groupValue = value;
        if (first)
            groupValue = data;
        if (data < groupValue)
            groupValue = data;

        return groupValue;
    }

    /**
     * Find max value from group records.
     * @param {Boolean} first - true if first record in set.
     * @param {Number} value - cumulative data from all previous group records.
     * @param {Number} data - data from current group record
     * @returns {Number} - max value from set.
     */
    static maxCase(first, value, data) {
        let groupValue = value;
        if (first)
            groupValue = data;
        if (data > groupValue)
            groupValue = data;

        return groupValue;
    }
}

/** Fields from all tables. */
class TableFields {
    constructor() {
        /** @property {TableField[]} */
        this.allFields = [];
        /** @property {Map<String, TableField>} */
        this.fieldNameMap = new Map();
        /** @property {Map<String, TableField>} */
        this.tableColumnMap = new Map();
    }

    /**
     * Iterate through all table fields and create a list of these VirtualFields.
     * @param {String} primaryTable - primary FROM table name in select.
     * @param {Map<String,Table>} tableInfo - map of all loaded tables. 
     */
    loadVirtualFields(primaryTable, tableInfo) {
        /** @type {String} */
        let tableName = "";
        /** @type {Table} */
        let tableObject = null;
        // @ts-ignore
        for ([tableName, tableObject] of tableInfo.entries()) {
            const validFieldNames = tableObject.getAllFieldNames();

            for (const field of validFieldNames) {
                const tableColumn = tableObject.getFieldColumn(field);
                if (tableColumn !== -1) {
                    let virtualField = this.findTableField(tableName, tableColumn);
                    if (virtualField !== null) {
                        virtualField.addAlias(field);
                    }
                    else {
                        virtualField = new TableField()
                            .setOriginalTable(tableName)
                            .setOriginalTableColumn(tableColumn)
                            .addAlias(field)
                            .setIsPrimaryTable(primaryTable.toUpperCase() === tableName.toUpperCase())
                            .setTableInfo(tableObject);

                        this.allFields.push(virtualField);
                    }

                    this.indexTableField(virtualField, primaryTable.toUpperCase() === tableName.toUpperCase());
                }
            }
        }

        this.allFields.sort(TableFields.sortPrimaryFields);
    }

    /**
     * Sort function for table fields list.
     * @param {TableField} fldA 
     * @param {TableField} fldB 
     */
    static sortPrimaryFields(fldA, fldB) {
        let keyA = fldA.isPrimaryTable ? 0 : 1000;
        let keyB = fldB.isPrimaryTable ? 0 : 1000;

        keyA += fldA.originalTableColumn;
        keyB += fldB.originalTableColumn;

        if (keyA < keyB)
            return -1;
        else if (keyA > keyB)
            return 1;
        return 0;
    }

    /**
     * Set up mapping to quickly find field info - by all (alias) names, by table+column.
     * @param {TableField} field - field info.
     * @param {Boolean} isPrimaryTable - is this a field from the SELECT FROM TABLE.
     */
    indexTableField(field, isPrimaryTable = false) {
        for (const aliasField of field.aliasNames) {
            const fieldInfo = this.fieldNameMap.get(aliasField.toUpperCase());

            if (typeof fieldInfo === 'undefined' || isPrimaryTable) {
                this.fieldNameMap.set(aliasField.toUpperCase(), field);
            }
        }

        //  This is something referenced in GROUP BY but is NOT in the SELECTED fields list.
        if (field.tempField && !this.fieldNameMap.has(field.columnName.toUpperCase())) {
            this.fieldNameMap.set(field.columnName.toUpperCase(), field);
        }

        if (field.originalTableColumn !== -1) {
            const key = `${field.originalTable}:${field.originalTableColumn}`;
            if (!this.tableColumnMap.has(key))
                this.tableColumnMap.set(key, field);
        }
    }

    /**
     * Quickly find field info for TABLE + COLUMN NUMBER (key of map)
     * @param {String} tableName - Table name to search for.
     * @param {Number} tableColumn - Column number to search for.
     * @returns {TableField} -located table info (null if not found).
     */
    findTableField(tableName, tableColumn) {
        const key = `${tableName}:${tableColumn}`;

        if (!this.tableColumnMap.has(key)) {
            return null;
        }

        return this.tableColumnMap.get(key);
    }

    /**
     * Is this field in our map.
     * @param {String} field - field name
     * @returns {Boolean} - found in map if true.
     */
    hasField(field) {
        return this.fieldNameMap.has(field.toUpperCase());
    }

    /**
     * Get field info.
     * @param {String} field - table column name to find 
     * @returns {TableField} - table info (undefined if not found)
     */
    getFieldInfo(field) {
        return this.fieldNameMap.get(field.toUpperCase());
    }

    /**
     * Get table associated with field name.
     * @param {String} field - field name to search for
     * @returns {Table} - associated table info (undefined if not found)
     */
    getTableInfo(field) {
        const fldInfo = this.getFieldInfo(field);

        return typeof fldInfo !== 'undefined' ? fldInfo.tableInfo : fldInfo;
    }

    /**
     * Get column number for field.
     * @param {String} field - field name
     * @returns {Number} - column number in table for field (-1 if not found)
     */
    getFieldColumn(field) {
        const fld = this.getFieldInfo(field);
        if (fld !== null) {
            return fld.tableColumn;
        }

        return -1;
    }

    /**
     * Get field column number.
     * @param {String} field - field name
     * @returns {Number} - column number.
     */
    getSelectFieldColumn(field) {
        let fld = this.getFieldInfo(field);
        if (typeof fld !== 'undefined' && fld.selectColumn !== -1) {
            return fld.selectColumn;
        }

        for (fld of this.getSelectFields()) {
            if (fld.aliasNames.indexOf(field.toUpperCase()) !== -1) {
                return fld.selectColumn;
            }
        }

        return -1;
    }

    /**
     * Updates internal SELECTED (returned in data) field list.
     * @param {Object} astFields - AST from SELECT
     * @param {Number} nextColumnPosition
     * @param {Boolean} isTempField
     */
    updateSelectFieldList(astFields, nextColumnPosition, isTempField) {
        for (const selField of astFields) {
            const parsedField = this.parseAstSelectField(selField);
            const columnTitle = (typeof selField.as !== 'undefined' && selField.as !== "" ? selField.as : selField.name);

            const selectedFieldParms = {
                selField, parsedField, columnTitle, nextColumnPosition, isTempField
            };

            if (parsedField.calculatedField === null && this.hasField(parsedField.columnName)) {
                this.updateColumnAsSelected(selectedFieldParms);
                nextColumnPosition = selectedFieldParms.nextColumnPosition;
            }
            else if (parsedField.calculatedField !== null) {
                this.updateCalculatedAsSelected(selectedFieldParms);
                nextColumnPosition++;
            }
            else {
                this.updateConstantAsSelected(selectedFieldParms);
                nextColumnPosition++;
            }           
        }
    }

    updateColumnAsSelected(selectedFieldParms) {
        let fieldInfo = this.getFieldInfo(selectedFieldParms.parsedField.columnName);

        //  If GROUP BY field is in our SELECT field list - we can ignore.
        if (selectedFieldParms.isTempField && fieldInfo.selectColumn !== -1)
            return;
        
        if (selectedFieldParms.parsedField.aggregateFunctionName !== "" || fieldInfo.selectColumn !== -1) {
            //  A new SELECT field, not from existing.
            const newFieldInfo = new TableField();
            Object.assign(newFieldInfo, fieldInfo);
            fieldInfo = newFieldInfo;

            this.allFields.push(fieldInfo);
        }

        fieldInfo
            .setAggregateFunction(selectedFieldParms.parsedField.aggregateFunctionName)
            .setColumnTitle(selectedFieldParms.columnTitle)
            .setColumnName(selectedFieldParms.selField.name)
            .setDistinctSetting(selectedFieldParms.parsedField.fieldDistinct)
            .setSelectColumn(selectedFieldParms.nextColumnPosition)
            .setIsTempField(selectedFieldParms.isTempField);

        selectedFieldParms.nextColumnPosition++;

        this.indexTableField(fieldInfo);
    }

    updateCalculatedAsSelected(selectedFieldParms) {
        const fieldInfo = new TableField();
        this.allFields.push(fieldInfo);

        fieldInfo
            .setColumnTitle(selectedFieldParms.columnTitle)
            .setColumnName(selectedFieldParms.selField.name)
            .setSelectColumn(selectedFieldParms.nextColumnPosition)
            .setCalculatedFormula(selectedFieldParms.selField.name)
            .setSubQueryAst(selectedFieldParms.selField.subQuery)
            .setIsTempField(selectedFieldParms.isTempField);

        this.indexTableField(fieldInfo);
    }

    updateConstantAsSelected(selectedFieldParms) {
        const fieldInfo = new TableField();
        this.allFields.push(fieldInfo);

        fieldInfo
            .setCalculatedFormula(selectedFieldParms.parsedField.columnName)
            .setAggregateFunction(selectedFieldParms.parsedField.aggregateFunctionName)
            .setSelectColumn(selectedFieldParms.nextColumnPosition)
            .setColumnName(selectedFieldParms.selField.name)
            .setColumnTitle(selectedFieldParms.columnTitle)
            .setIsTempField(selectedFieldParms.isTempField);;

        this.indexTableField(fieldInfo);
    }

    /**
     * Fields in GROUP BY and ORDER BY might not be in the SELECT field list.  Add a TEMP version to that list.
     * @param {Object} ast - AST to search for GROUP BY and ORDER BY.
     */
    addReferencedColumnstoSelectFieldList(ast) {
        this.addTempMissingSelectedField(ast['ORDER BY']);
    }

    /**
     * Add to Select field list as a temporary field for the fields in AST.
     * @param {Object} astColumns - find columns mentioned not already in Select Field List
     */
    addTempMissingSelectedField(astColumns) {
        if (typeof astColumns !== 'undefined') {
            for (const order of astColumns) {
                if (this.getSelectFieldColumn(order.name) === -1) {
                    const fieldInfo = this.getFieldInfo(order.name);

                    //  A new SELECT field, not from existing.
                    const newFieldInfo = new TableField();
                    Object.assign(newFieldInfo, fieldInfo);
                    newFieldInfo
                        .setSelectColumn(this.getNextSelectColumnNumber())
                        .setIsTempField(true);

                    this.allFields.push(newFieldInfo);
                }
            }
        }
    }

    /**
     * Find next available column number in selected field list.
     * @returns {Number} - column number
     */
    getNextSelectColumnNumber() {
        let next = -1;
        for (const fld of this.getSelectFields()) {
            next = fld.selectColumn > next ? fld.selectColumn : next;
        }

        return next === -1 ? next : ++next;
    }

    /**
     * Return a list of temporary column numbers in select field list.
     * @returns {Number[]} - sorted list of temp column numbers.
     */
    getTempSelectedColumnNumbers() {
        /** @type {Number[]} */
        const tempCols = [];
        for (const fld of this.getSelectFields()) {
            if (fld.tempField) {
                tempCols.push(fld.selectColumn);
            }
        }
        tempCols.sort((a, b) => (b - a));

        return tempCols;
    }

    /**
     * Get a sorted list (by column number) of selected fields.
     * @returns {TableField[]} - selected fields
     */
    getSelectFields() {
        const selectedFields = this.allFields.filter((a) => a.selectColumn !== -1);
        selectedFields.sort((a, b) => a.selectColumn - b.selectColumn);

        return selectedFields;
    }

    /**
     * Get SELECTED Field names sorted list of column number.
     * @returns {String[]} - Table field names
     */
    getColumnNames() {
        const columnNames = [];

        for (const fld of this.getSelectFields()) {
            columnNames.push(fld.columnName);
        }

        return columnNames;
    }

    /**
     * Get column titles. If alias was set, that column would be the alias, otherwise it is column name.
     * @param {String} columnTableNameReplacement
     * @returns {String[]} - column titles
     */
    getColumnTitles(columnTableNameReplacement) {
        const columnTitles = [];

        for (const fld of this.getSelectFields()) {
            if (!fld.tempField) {
                let columnOutput = fld.columnTitle;

                //  When subquery table data becomes data for the derived table name, references to
                //  original table names in column output needs to be changed to new derived table name.
                if (columnTableNameReplacement !== null && columnOutput.startsWith(`${fld.originalTable}.`)) {
                    columnOutput = columnOutput.replace(`${fld.originalTable}.`, `${columnTableNameReplacement}.`);
                }
                columnTitles.push(columnOutput);
            }
        }

        return columnTitles;
    }

    /**
     * Derived tables will cause an update to any TableField.  It updates with a new column number and new table (derived) info.
     * @param {DerivedTable} derivedTable - derived table info.
     */
    updateDerivedTableVirtualFields(derivedTable) {
        const derivedTableFields = derivedTable.tableInfo.getAllVirtualFields();

        let fieldNo = 0;
        for (const field of derivedTableFields) {
            if (this.hasField(field.fieldName)) {
                const originalField = this.getFieldInfo(field.fieldName);
                originalField.derivedTableColumn = fieldNo;
                originalField.tableInfo = derivedTable.tableInfo;
            }

            fieldNo++;
        }
    }

    /**
     * @typedef {Object} ParsedSelectField
     * @property {String} columnName
     * @property {String} aggregateFunctionName
     * @property {Object} calculatedField
     * @property {String} fieldDistinct
     */

    /**
     * Parse SELECT field in AST (may include functions or calculations)
     * @param {Object} selField 
     * @returns {ParsedSelectField}
     */
    parseAstSelectField(selField) {
        let columnName = selField.name;
        let aggregateFunctionName = "";
        let fieldDistinct = "";
        const calculatedField = (typeof selField.terms === 'undefined') ? null : selField.terms;

        if (calculatedField === null && !this.hasField(columnName)) {
            const functionNameRegex = /^\w+\s*(?=\()/;
            let matches = columnName.match(functionNameRegex)
            if (matches !== null && matches.length > 0)
                aggregateFunctionName = matches[0].trim();

            matches = SelectTables.parseForFunctions(columnName, aggregateFunctionName);
            if (matches !== null && matches.length > 1) {
                columnName = matches[1];

                // e.g.  select count(distinct field)    OR   select count(all field)
                [columnName, fieldDistinct] = TableFields.getSelectCountModifiers(columnName);
            }
        }

        return { columnName, aggregateFunctionName, calculatedField, fieldDistinct };
    }

    /**
     * Parse for any SELECT COUNT modifiers like 'DISTINCT' or 'ALL'.
     * @param {String} originalColumnName - column (e.g. 'distinct customer_id')
     * @returns {String[]} - [0] - parsed column name, [1] - count modifier
     */
    static getSelectCountModifiers(originalColumnName) {
        let fieldDistinct = "";
        let columnName = originalColumnName;

        //  e.g.  count(distinct field)
        const distinctParts = columnName.split(" ");
        if (distinctParts.length > 1) {
            const distinctModifiers = ["DISTINCT", "ALL"];
            if (distinctModifiers.includes(distinctParts[0].toUpperCase())) {
                fieldDistinct = distinctParts[0].toUpperCase();
                columnName = distinctParts[1];
            }
        }

        return [columnName, fieldDistinct];
    }

    /**
     * Counts the number of conglomerate field functions in SELECT field list.
     * @returns {Number} - Number of conglomerate functions.
     */
    getConglomerateFieldCount() {
        let count = 0;
        for (/** @type {TableField} */ const field of this.getSelectFields()) {
            if (field.aggregateFunction !== "")
                count++;
        }

        return count;
    }
}

/** Table column information. */
class TableField {
    constructor() {
        /** @property {String} */
        this.originalTable = "";
        /** @property {Number} */
        this.originalTableColumn = -1;
        /** @property {String[]} */
        this.aliasNames = [];
        /** @property {String} */
        this.fieldName = "";
        /** @property {Number} */
        this.derivedTableColumn = -1;
        /** @property {Number} */
        this.selectColumn = -1;
        /** @property {Boolean} */
        this.tempField = false;
        /** @property {String} */
        this.calculatedFormula = "";
        /** @property {String} */
        this.aggregateFunction = "";
        /** @property {String} */
        this.columnTitle = "";
        /** @property {String} */
        this.columnName = "";
        /** @property {String} */
        this.distinctSetting = "";
        /** @property {Object} */
        this.subQueryAst = null;
        /** @property {Boolean} */
        this._isPrimaryTable = false;
        /** @property {Table} */
        this.tableInfo = null;
    }

    /**
     * Get field column number.
     * @returns {Number} - column number
     */
    get tableColumn() {
        return this.derivedTableColumn === -1 ? this.originalTableColumn : this.derivedTableColumn;
    }

    /**
     * Original table name before any derived table updates.
     * @param {String} table - original table name
     * @returns {TableField}
     */
    setOriginalTable(table) {
        this.originalTable = table.trim().toUpperCase();
        return this;
    }

    /**
     * Column name found in column title row.
     * @param {Number} column 
     * @returns {TableField}
     */
    setOriginalTableColumn(column) {
        this.originalTableColumn = column;
        return this;
    }

    /**
     * Alias name assigned to field in select statement.
     * @param {String} columnAlias - alias name
     * @returns {TableField}
     */
    addAlias(columnAlias) {
        const alias = columnAlias.trim().toUpperCase();
        if (this.fieldName === "" || alias.indexOf(".") !== -1) {
            this.fieldName = alias;
        }

        if (this.aliasNames.indexOf(alias) === -1) {
            this.aliasNames.push(alias);
        }

        return this;
    }

    /**
     * Set column number in table data for field.
     * @param {Number} column - column number.
     * @returns {TableField}
     */
    setSelectColumn(column) {
        this.selectColumn = column;

        return this;
    }

    /**
     * Fields referenced BUT not in final output.
     * @param {Boolean} value 
     * @returns {TableField}
     */
    setIsTempField(value) {
        this.tempField = value;
        return this;
    }

    /**
     * Aggregate function number used (e.g. 'SUM')
     * @param {String} value - aggregate function name or ''
     * @returns {TableField}
     */
    setAggregateFunction(value) {
        this.aggregateFunction = value.toUpperCase();
        return this;
    }

    /**
     * Calculated formula for field (e.g. 'CASE WHEN QUANTITY >= 100 THEN 1 ELSE 0 END')
     * @param {String} value 
     * @returns {TableField}
     */
    setCalculatedFormula(value) {
        this.calculatedFormula = value;
        return this;
    }

    /**
     * The AST from just the subquery in the SELECT.
     * @param {Object} ast - subquery ast.
     * @returns {TableField}
     */
    setSubQueryAst(ast) {
        this.subQueryAst = ast;
        return this;
    }

    /**
     * Set column TITLE.  If an alias is available, that is used - otherwise it is column name.
     * @param {String} columnTitle - column title used in output
     * @returns {TableField}
     */
    setColumnTitle(columnTitle) {
        this.columnTitle = columnTitle;
        return this;
    }

    /**
     * Set the columnname.
     * @param {String} columnName 
     * @returns {TableField}
     */
    setColumnName(columnName) {
        this.columnName = columnName;
        return this;
    }

    /**
     * Set any count modified like 'DISTINCT' or 'ALL'.
     * @param {String} distinctSetting 
     * @returns {TableField}
     */
    setDistinctSetting(distinctSetting) {
        this.distinctSetting = distinctSetting;
        return this
    }

    /**
     * Set if this field belongs to primary table (i.e. select * from table), rather than a joined tabled.
     * @param {Boolean} isPrimary - true if from primary table.
     * @returns {TableField}
     */
    setIsPrimaryTable(isPrimary) {
        this._isPrimaryTable = isPrimary;
        return this;
    }

    /**
     * Is this field in the primary table.
     * @returns {Boolean}
     */
    get isPrimaryTable() {
        return this._isPrimaryTable;
    }

    /**
     * Link this field to the table info.
     * @param {Table} tableInfo 
     * @returns {TableField}
     */
    setTableInfo(tableInfo) {
        this.tableInfo = tableInfo;
        return this;
    }

    /**
     * Retrieve field data for tableRow
     * @param {Number} tableRow - row to read data from
     * @returns {any} - data
     */
    getData(tableRow) {
        const columnNumber = this.derivedTableColumn === -1 ? this.originalTableColumn : this.derivedTableColumn;
        if (tableRow < 0 || columnNumber < 0)
            return "";

        return this.tableInfo.tableData[tableRow][columnNumber];
    }

    /**
     * Search through list of fields and return a list of those that include the table name (e.g. TABLE.COLUMN vs COLUMN)
     * @param {TableField[]} masterFields 
     * @returns {String[]}
     */
    static getAllExtendedAliasNames(masterFields) {
        const concatFields = [];
        for (const vField of masterFields) {
            for (const aliasName of vField.aliasNames) {
                if (aliasName.indexOf(".") !== -1) {
                    concatFields.push(aliasName);
                }
            }
        }

        return concatFields;
    }
}


//  Code inspired from:  https://github.com/dsferruzza/simpleSqlParser

/** Parse SQL SELECT statement and convert into Abstract Syntax Tree */
class SqlParse {
    /**
     * 
     * @param {String} cond 
     * @returns {String}
     */
    static sqlCondition2JsCondition(cond) {
        const ast = SqlParse.sql2ast(`SELECT A FROM c WHERE ${cond}`);
        let sqlData = "";

        if (typeof ast.WHERE !== 'undefined') {
            const conditions = ast.WHERE;
            if (typeof conditions.logic === 'undefined')
                sqlData = SqlParse.resolveSqlCondition("OR", [conditions]);
            else
                sqlData = SqlParse.resolveSqlCondition(conditions.logic, conditions.terms);

        }

        return sqlData;
    }

    /**
     * Parse a query
     * @param {String} query 
     * @returns {Object}
     */
    static sql2ast(query) {
        // Define which words can act as separator
        const myKeyWords = SqlParse.generateUsedKeywordList(query);
        const [parts_name, parts_name_escaped] = SqlParse.generateSqlSeparatorWords(myKeyWords);

        //  Include brackets around separate selects used in things like UNION, INTERSECT...
        let modifiedQuery = SqlParse.sqlStatementSplitter(query);

        // Hide words defined as separator but written inside brackets in the query
        modifiedQuery = SqlParse.hideInnerSql(modifiedQuery, parts_name_escaped, SqlParse.protect);

        // Write the position(s) in query of these separators
        const parts_order = SqlParse.getPositionsOfSqlParts(modifiedQuery, parts_name);

        // Delete duplicates (caused, for example, by JOIN and INNER JOIN)
        SqlParse.removeDuplicateEntries(parts_order);

        // Generate protected word list to reverse the use of protect()
        let words = parts_name_escaped.slice(0);
        words = words.map(item => SqlParse.protect(item));

        // Split parts
        const parts = modifiedQuery.split(new RegExp(parts_name_escaped.join('|'), 'i'));

        // Unhide words precedently hidden with protect()
        for (let i = 0; i < parts.length; i++) {
            parts[i] = SqlParse.hideInnerSql(parts[i], words, SqlParse.unprotect);
        }

        // Analyze parts
        const result = SqlParse.analyzeParts(parts_order, parts);

        if (typeof result.FROM !== 'undefined' && typeof result.FROM.FROM !== 'undefined' && typeof result.FROM.FROM.as !== 'undefined' && result.FROM.FROM.as !== '') {
            //   Subquery FROM creates an ALIAS name, which is then used as FROM table name.
            result.FROM.table = result.FROM.FROM.as;
            result.FROM.isDerived = true;
        }

        return result;
    }

    /**
    * 
    * @param {String} logic 
    * @param {Object} terms 
    * @returns {String}
    */
    static resolveSqlCondition(logic, terms) {
        let jsCondition = "";

        for (const cond of terms) {
            if (typeof cond.logic === 'undefined') {
                if (jsCondition !== "" && logic === "AND") {
                    jsCondition += " && ";
                }
                else if (jsCondition !== "" && logic === "OR") {
                    jsCondition += " || ";
                }

                jsCondition += ` ${cond.left}`;
                if (cond.operator === "=")
                    jsCondition += " == ";
                else
                    jsCondition += ` ${cond.operator}`;
                jsCondition += ` ${cond.right}`;
            }
            else {
                jsCondition += SqlParse.resolveSqlCondition(cond.logic, cond.terms);
            }
        }

        return jsCondition;
    }

    /**
     * 
     * @param {String} query
     * @returns {String[]} 
     */
    static generateUsedKeywordList(query) {
        const generatedList = new Set();
        // Define which words can act as separator
        const keywords = ['SELECT', 'FROM', 'JOIN', 'LEFT JOIN', 'RIGHT JOIN', 'INNER JOIN', 'FULL JOIN', 'ORDER BY', 'GROUP BY', 'HAVING', 'WHERE', 'LIMIT', 'UNION ALL', 'UNION', 'INTERSECT', 'EXCEPT', 'PIVOT'];

        const modifiedQuery = query.toUpperCase();

        for (const word of keywords) {
            let pos = 0;
            while (pos !== -1) {
                pos = modifiedQuery.indexOf(word, pos);

                if (pos !== -1) {
                    generatedList.add(query.substring(pos, pos + word.length));
                    pos++;
                }
            }
        }

        // @ts-ignore
        return [...generatedList];
    }

    /**
     * 
     * @param {String[]} keywords 
     * @returns {String[][]}
     */
    static generateSqlSeparatorWords(keywords) {
        let parts_name = keywords.map(item => `${item} `);
        parts_name = parts_name.concat(keywords.map(item => `${item}(`));
        const parts_name_escaped = parts_name.map(item => item.replace('(', '[\\(]'));

        return [parts_name, parts_name_escaped];
    }

    /**
     * 
     * @param {String} src 
     * @returns {String}
     */
    static sqlStatementSplitter(src) {
        let newStr = src;

        // Define which words can act as separator
        const reg = SqlParse.makeSqlPartsSplitterRegEx(["UNION ALL", "UNION", "INTERSECT", "EXCEPT"]);

        const matchedUnions = newStr.match(reg);
        if (matchedUnions === null || matchedUnions.length === 0)
            return newStr;

        let prefix = "";
        const parts = [];
        let pos = newStr.search(matchedUnions[0]);
        if (pos > 0) {
            prefix = newStr.substring(0, pos);
            newStr = newStr.substring(pos + matchedUnions[0].length);
        }

        for (let i = 1; i < matchedUnions.length; i++) {
            const match = matchedUnions[i];
            pos = newStr.search(match);

            parts.push(newStr.substring(0, pos));
            newStr = newStr.substring(pos + match.length);
        }
        if (newStr.length > 0)
            parts.push(newStr);

        newStr = prefix;
        for (let i = 0; i < matchedUnions.length; i++) {
            newStr += `${matchedUnions[i]} (${parts[i]}) `;
        }

        return newStr;
    }

    /**
     * 
     * @param {String[]} keywords 
     * @returns {RegExp}
     */
    static makeSqlPartsSplitterRegEx(keywords) {
        // Define which words can act as separator
        let parts_name = keywords.map(item => `${item} `);
        parts_name = parts_name.concat(keywords.map(item => `${item}(`));
        parts_name = parts_name.concat(parts_name.map(item => item.toLowerCase()));
        const parts_name_escaped = parts_name.map(item => item.replace('(', '[\\(]'));

        return new RegExp(parts_name_escaped.join('|'), 'gi');
    }

    /**
     * 
     * @param {String} str 
     * @param {String[]} parts_name_escaped
     * @param {Object} replaceFunction
     */
    static hideInnerSql(str, parts_name_escaped, replaceFunction) {
        if (str.indexOf("(") === -1 && str.indexOf(")") === -1)
            return str;

        let bracketCount = 0;
        let endCount = -1;
        let newStr = str;

        for (let i = newStr.length - 1; i >= 0; i--) {
            const ch = newStr.charAt(i);

            if (ch === ")") {
                bracketCount++;

                if (bracketCount === 1) {
                    endCount = i;
                }
            }
            else if (ch === "(") {
                bracketCount--;
                if (bracketCount === 0) {

                    let query = newStr.substring(i, endCount + 1);

                    // Hide words defined as separator but written inside brackets in the query
                    query = query.replace(new RegExp(parts_name_escaped.join('|'), 'gi'), replaceFunction);

                    newStr = newStr.substring(0, i) + query + newStr.substring(endCount + 1);
                }
            }
        }
        return newStr;
    }

    /**
     * 
     * @param {String} modifiedQuery 
     * @param {String[]} parts_name 
     * @returns {String[]}
     */
    static getPositionsOfSqlParts(modifiedQuery, parts_name) {
        // Write the position(s) in query of these separators
        const parts_order = [];
        function realNameCallback(_match, name) {
            return name;
        }
        parts_name.forEach(function (item) {
            let pos = 0;
            let part = 0;

            do {
                part = modifiedQuery.indexOf(item, pos);
                if (part !== -1) {
                    const realName = item.replace(/^((\w|\s)+?)\s?\(?$/i, realNameCallback);

                    if (typeof parts_order[part] === 'undefined' || parts_order[part].length < realName.length) {
                        parts_order[part] = realName;	// Position won't be exact because the use of protect()  (above) and unprotect() alter the query string ; but we just need the order :)
                    }

                    pos = part + realName.length;
                }
            }
            while (part !== -1);
        });

        return parts_order;
    }

    /**
     * Delete duplicates (caused, for example, by JOIN and INNER JOIN)
     * @param {String[]} parts_order
     */
    static removeDuplicateEntries(parts_order) {
        let busy_until = 0;
        parts_order.forEach(function (item, key) {
            if (busy_until > key)
                delete parts_order[key];
            else {
                busy_until = key + item.length;

                // Replace JOIN by INNER JOIN
                if (item.toUpperCase() === 'JOIN')
                    parts_order[key] = 'INNER JOIN';
            }
        });
    }

    /**
     * Add some # inside a string to avoid it to match a regex/split
     * @param {String} str 
     * @returns {String}
     */
    static protect(str) {
        let result = '#';
        const length = str.length;
        for (let i = 0; i < length; i++) {
            result += `${str[i]}#`;
        }
        return result;
    }

    /**
     * Restore a string output by protect() to its original state
     * @param {String} str 
     * @returns {String}
     */
    static unprotect(str) {
        let result = '';
        const length = str.length;
        for (let i = 1; i < length; i = i + 2) result += str[i];
        return result;
    }

    /**
     * 
     * @param {String[]} parts_order 
     * @param {String[]} parts 
     * @returns {Object}
     */
    static analyzeParts(parts_order, parts) {
        const result = {};
        let j = 0;
        parts_order.forEach(function (item, _key) {
            const itemName = item.toUpperCase();
            j++;
            const part_result = SelectKeywordAnalysis.analyze(item, parts[j]);

            if (typeof result[itemName] !== 'undefined') {
                if (typeof result[itemName] === 'string' || typeof result[itemName][0] === 'undefined') {
                    const tmp = result[itemName];
                    result[itemName] = [];
                    result[itemName].push(tmp);
                }

                result[itemName].push(part_result);
            }
            else {
                result[itemName] = part_result;
            }

        });

        // Reorganize joins
        SqlParse.reorganizeJoins(result);

        if (typeof result.JOIN !== 'undefined') {
            result.JOIN.forEach((item, key) => result.JOIN[key].cond = CondParser.parse(item.cond));
        }

        SqlParse.reorganizeUnions(result);

        return result;
    }

    /**
     * 
     * @param {Object} result 
     */
    static reorganizeJoins(result) {
        const joinArr = [
            ['FULL JOIN', 'full'],
            ['RIGHT JOIN', 'right'],
            ['INNER JOIN', 'inner'],
            ['LEFT JOIN', 'left']
        ];

        for (const join of joinArr) {
            const [joinName, joinType] = join;
            SqlParse.reorganizeSpecificJoin(result, joinName, joinType);
        }
    }

    /**
     * 
     * @param {Object} result 
     * @param {String} joinName 
     * @param {String} joinType 
     */
    static reorganizeSpecificJoin(result, joinName, joinType) {
        if (typeof result[joinName] !== 'undefined') {
            if (typeof result.JOIN === 'undefined') result.JOIN = [];
            if (typeof result[joinName][0] !== 'undefined') {
                result[joinName].forEach(function (item) {
                    item.type = joinType;
                    result.JOIN.push(item);
                });
            }
            else {
                result[joinName].type = joinType;
                result.JOIN.push(result[joinName]);
            }
            delete result[joinName];
        }
    }

    /**
     * 
     * @param {Object} result 
     */
    static reorganizeUnions(result) {
        const astRecursiveTableBlocks = ['UNION', 'UNION ALL', 'INTERSECT', 'EXCEPT'];

        for (const union of astRecursiveTableBlocks) {
            if (typeof result[union] === 'string') {
                result[union] = [SqlParse.sql2ast(SqlParse.parseUnion(result[union]))];
            }
            else if (typeof result[union] !== 'undefined') {
                for (let i = 0; i < result[union].length; i++) {
                    result[union][i] = SqlParse.sql2ast(SqlParse.parseUnion(result[union][i]));
                }
            }
        }
    }

    static parseUnion(inStr) {
        let unionString = inStr;
        if (unionString.startsWith("(") && unionString.endsWith(")")) {
            unionString = unionString.substring(1, unionString.length - 1);
        }

        return unionString;
    }
}

/*
 * LEXER & PARSER FOR SQL CONDITIONS
 * Inspired by https://github.com/DmitrySoshnikov/Essentials-of-interpretation
 */

/** Lexical analyzer for SELECT statement. */
class CondLexer {
    constructor(source) {
        this.source = source;
        this.cursor = 0;
        this.currentChar = "";
        this.startQuote = "";
        this.bracketCount = 0;

        this.readNextChar();
    }

    // Read the next character (or return an empty string if cursor is at the end of the source)
    readNextChar() {
        if (typeof this.source !== 'string') {
            this.currentChar = "";
        }
        else {
            this.currentChar = this.source[this.cursor++] || "";
        }
    }

    // Determine the next token
    readNextToken() {
        if (/\w/.test(this.currentChar))
            return this.readWord();
        if (/["'`]/.test(this.currentChar))
            return this.readString();
        if (/[()]/.test(this.currentChar))
            return this.readGroupSymbol();
        if (/[!=<>]/.test(this.currentChar))
            return this.readOperator();
        if (/[+\-*/%]/.test(this.currentChar))
            return this.readMathOperator();
        if (this.currentChar === '?')
            return this.readBindVariable();

        if (this.currentChar === "") {
            return { type: 'eot', value: '' };
        }

        this.readNextChar();
        return { type: 'empty', value: '' };
    }

    readWord() {
        let tokenValue = "";
        this.bracketCount = 0;
        let insideQuotedString = false;
        this.startQuote = "";

        while (/./.test(this.currentChar)) {
            // Check if we are in a string
            insideQuotedString = this.isStartOrEndOfString(insideQuotedString);

            if (this.isFinishedWord(insideQuotedString))
                break;

            tokenValue += this.currentChar;
            this.readNextChar();
        }

        if (/^(AND|OR)$/i.test(tokenValue)) {
            return { type: 'logic', value: tokenValue.toUpperCase() };
        }

        if (/^(IN|IS|NOT|LIKE|NOT EXISTS|EXISTS)$/i.test(tokenValue)) {
            return { type: 'operator', value: tokenValue.toUpperCase() };
        }

        return { type: 'word', value: tokenValue };
    }

    /**
     * 
     * @param {Boolean} insideQuotedString 
     * @returns {Boolean}
     */
    isStartOrEndOfString(insideQuotedString) {
        if (!insideQuotedString && /['"`]/.test(this.currentChar)) {
            this.startQuote = this.currentChar;

            return true;
        }
        else if (insideQuotedString && this.currentChar === this.startQuote) {
            //  End of quoted string.
            return false;
        }

        return insideQuotedString;
    }

    /**
     * 
     * @param {Boolean} insideQuotedString 
     * @returns {Boolean}
     */
    isFinishedWord(insideQuotedString) {
        if (insideQuotedString)
            return false;

        // Token is finished if there is a closing bracket outside a string and with no opening
        if (this.currentChar === ')' && this.bracketCount <= 0) {
            return true;
        }

        if (this.currentChar === '(') {
            this.bracketCount++;
        }
        else if (this.currentChar === ')') {
            this.bracketCount--;
        }

        // Token is finished if there is a operator symbol outside a string
        if (/[!=<>]/.test(this.currentChar)) {
            return true;
        }

        // Token is finished on the first space which is outside a string or a function
        return this.currentChar === ' ' && this.bracketCount <= 0;
    }

    readString() {
        let tokenValue = "";
        const quote = this.currentChar;

        tokenValue += this.currentChar;
        this.readNextChar();

        while (this.currentChar !== quote && this.currentChar !== "") {
            tokenValue += this.currentChar;
            this.readNextChar();
        }

        tokenValue += this.currentChar;
        this.readNextChar();

        // Handle this case : `table`.`column`
        if (this.currentChar === '.') {
            tokenValue += this.currentChar;
            this.readNextChar();
            tokenValue += this.readString().value;

            return { type: 'word', value: tokenValue };
        }

        return { type: 'string', value: tokenValue };
    }

    readGroupSymbol() {
        const tokenValue = this.currentChar;
        this.readNextChar();

        return { type: 'group', value: tokenValue };
    }

    readOperator() {
        let tokenValue = this.currentChar;
        this.readNextChar();

        if (/[=<>]/.test(this.currentChar)) {
            tokenValue += this.currentChar;
            this.readNextChar();
        }

        return { type: 'operator', value: tokenValue };
    }

    readMathOperator() {
        const tokenValue = this.currentChar;
        this.readNextChar();

        return { type: 'mathoperator', value: tokenValue };
    }

    readBindVariable() {
        let tokenValue = this.currentChar;
        this.readNextChar();

        while (/\d/.test(this.currentChar)) {
            tokenValue += this.currentChar;
            this.readNextChar();
        }

        return { type: 'bindVariable', value: tokenValue };
    }
}

/** SQL Condition parser class. */
class CondParser {
    constructor(source) {
        this.lexer = new CondLexer(source);
        this.currentToken = {};

        this.readNextToken();
    }

    // Parse a string
    static parse(source) {
        return new CondParser(source).parseExpressionsRecursively();
    }

    // Read the next token (skip empty tokens)
    readNextToken() {
        this.currentToken = this.lexer.readNextToken();
        while (this.currentToken.type === 'empty')
            this.currentToken = this.lexer.readNextToken();
        return this.currentToken;
    }

    // Wrapper function ; parse the source
    parseExpressionsRecursively() {
        return this.parseLogicalExpression();
    }

    // Parse logical expressions (AND/OR)
    parseLogicalExpression() {
        let leftNode = this.parseConditionExpression();

        while (this.currentToken.type === 'logic') {
            const logic = this.currentToken.value;
            this.readNextToken();

            const rightNode = this.parseConditionExpression();

            // If we are chaining the same logical operator, add nodes to existing object instead of creating another one
            if (typeof leftNode.logic !== 'undefined' && leftNode.logic === logic && typeof leftNode.terms !== 'undefined')
                leftNode.terms.push(rightNode);
            else {
                const terms = [leftNode, rightNode];
                leftNode = { 'logic': logic, 'terms': terms.slice(0) };
            }
        }

        return leftNode;
    }

    // Parse conditions ([word/string] [operator] [word/string])
    parseConditionExpression() {
        let left = this.parseBaseExpression();

        if (this.currentToken.type !== 'operator') {
            return left;
        }

        let operator = this.currentToken.value;
        this.readNextToken();

        // If there are 2 adjacent operators, join them with a space (exemple: IS NOT)
        if (this.currentToken.type === 'operator') {
            operator += ` ${this.currentToken.value}`;
            this.readNextToken();
        }

        let right = null;
        if (this.currentToken.type === 'group' && (operator === 'EXISTS' || operator === 'NOT EXISTS')) {
            [left, right] = this.parseSelectExistsSubQuery();
        } else {
            right = this.parseBaseExpression(operator);
        }

        return { operator, left, right };
    }

    /**
     * 
     * @returns {Object[]}
     */
    parseSelectExistsSubQuery() {
        let rightNode = null;
        const leftNode = '""';

        this.readNextToken();
        if (this.currentToken.type === 'word' && this.currentToken.value === 'SELECT') {
            rightNode = this.parseSelectIn("", true);
            if (this.currentToken.type === 'group') {
                this.readNextToken();
            }
        }

        return [leftNode, rightNode];
    }

    // Parse base items
    /**
     * 
     * @param {String} operator 
     * @returns {Object}
     */
    parseBaseExpression(operator = "") {
        let astNode = {};

        // If this is a word/string, return its value
        if (this.currentToken.type === 'word' || this.currentToken.type === 'string') {
            astNode = this.parseWordExpression();
        }
        // If this is a group, skip brackets and parse the inside
        else if (this.currentToken.type === 'group') {
            astNode = this.parseGroupExpression(operator);
        }
        else if (this.currentToken.type === 'bindVariable') {
            astNode = this.currentToken.value;
            this.readNextToken();
        }

        return astNode;
    }

    /**
     * 
     * @returns {Object}
     */
    parseWordExpression() {
        let astNode = this.currentToken.value;
        this.readNextToken();

        if (this.currentToken.type === 'mathoperator') {
            astNode += ` ${this.currentToken.value}`;
            this.readNextToken();
            while ((this.currentToken.type === 'mathoperator' || this.currentToken.type === 'word') && this.currentToken.type !== 'eot') {
                astNode += ` ${this.currentToken.value}`;
                this.readNextToken();
            }
        }

        return astNode;
    }

    /**
     * 
     * @param {String} operator 
     * @returns {Object}
     */
    parseGroupExpression(operator) {
        this.readNextToken();
        let astNode = this.parseExpressionsRecursively();

        const isSelectStatement = typeof astNode === "string" && astNode.toUpperCase() === 'SELECT';

        if (operator === 'IN' || isSelectStatement) {
            astNode = this.parseSelectIn(astNode, isSelectStatement);
        }
        else {
            //  Are we within brackets of mathmatical expression ?
            let inCurrentToken = this.currentToken;

            while (inCurrentToken.type !== 'group' && inCurrentToken.type !== 'eot') {
                this.readNextToken();
                if (inCurrentToken.type !== 'group') {
                    astNode += ` ${inCurrentToken.value}`;
                }

                inCurrentToken = this.currentToken;
            }

        }

        this.readNextToken();

        return astNode;
    }

    /**
     * 
     * @param {Object} startAstNode 
     * @param {Boolean} isSelectStatement 
     * @returns {Object}
     */
    parseSelectIn(startAstNode, isSelectStatement) {
        let astNode = startAstNode;
        let inCurrentToken = this.currentToken;
        let bracketCount = 1;
        while (bracketCount !== 0 && inCurrentToken.type !== 'eot') {
            this.readNextToken();
            if (isSelectStatement) {
                astNode += ` ${inCurrentToken.value}`;
            }
            else {
                astNode += `, ${inCurrentToken.value}`;
            }

            inCurrentToken = this.currentToken;
            bracketCount += CondParser.groupBracketIncrementer(inCurrentToken);
        }

        if (isSelectStatement) {
            astNode = SqlParse.sql2ast(astNode);
        }

        return astNode;
    }

    static groupBracketIncrementer(inCurrentToken) {
        let diff = 0;
        if (inCurrentToken.type === 'group') {
            if (inCurrentToken.value === '(') {
                diff = 1;
            }
            else if (inCurrentToken.value === ')') {
                diff = -1;
            }
        }

        return diff
    }
}

/** Analyze each distinct component of SELECT statement. */
class SelectKeywordAnalysis {
    static analyze(itemName, part) {
        const keyWord = itemName.toUpperCase().replace(/ /g, '_');

        if (typeof SelectKeywordAnalysis[keyWord] === 'undefined') {
            throw new Error(`Can't analyze statement ${itemName}`);
        }

        return SelectKeywordAnalysis[keyWord](part);
    }

    static SELECT(str, isOrderBy = false) {
        const selectParts = SelectKeywordAnalysis.protect_split(',', str);
        const selectResult = selectParts.filter(function (item) {
            return item !== '';
        }).map(function (item) {
            let order = "";
            if (isOrderBy) {
                const order_by = /^(.+?)(\s+ASC|DESC)?$/gi;
                const orderData = order_by.exec(item);
                if (orderData !== null) {
                    order = typeof orderData[2] === 'undefined' ? "ASC" : SelectKeywordAnalysis.trim(orderData[2]);
                    item = orderData[1].trim();
                }
            }

            //  Is there a column alias?
            const [name, as] = SelectKeywordAnalysis.getNameAndAlias(item);

            const splitPattern = /[\s()*/%+-]+/g;
            let terms = name.split(splitPattern);

            if (terms !== null) {
                const aggFunc = ["SUM", "MIN", "MAX", "COUNT", "AVG", "DISTINCT"];
                terms = (aggFunc.indexOf(terms[0].toUpperCase()) === -1) ? terms : null;
            }
            if (name !== "*" && terms !== null && terms.length > 1) {
                const subQuery = SelectKeywordAnalysis.parseForCorrelatedSubQuery(item);
                return { name, terms, as, subQuery, order };
            }

            return { name, as, order };
        });

        return selectResult;
    }

    static FROM(str) {
        const subqueryAst = this.parseForCorrelatedSubQuery(str);
        if (subqueryAst !== null) {
            //  If there is a subquery creating a DERIVED table, it must have a derived table name.
            //  Extract this subquery AS tableName.
            const [, alias] = SelectKeywordAnalysis.getNameAndAlias(str);
            if (alias !== "" && typeof subqueryAst.FROM !== 'undefined') {
                subqueryAst.FROM.as = alias.toUpperCase();
            }

            return subqueryAst;
        }

        let fromResult = str.split(',');

        fromResult = fromResult.map(item => SelectKeywordAnalysis.trim(item));
        fromResult = fromResult.map(item => {
            const [table, as] = SelectKeywordAnalysis.getNameAndAlias(item);
            return { table, as };
        });
        return fromResult[0];
    }

    static LEFT_JOIN(str) {
        return SelectKeywordAnalysis.allJoins(str);
    }

    static INNER_JOIN(str) {
        return SelectKeywordAnalysis.allJoins(str);
    }

    static RIGHT_JOIN(str) {
        return SelectKeywordAnalysis.allJoins(str);
    }

    static FULL_JOIN(str) {
        return SelectKeywordAnalysis.allJoins(str);
    }

    static allJoins(str) {
        const subqueryAst = this.parseForCorrelatedSubQuery(str);

        const strParts = str.toUpperCase().split(' ON ');
        const table = strParts[0].split(' AS ');
        const joinResult = {};
        joinResult.table = subqueryAst !== null ? subqueryAst : SelectKeywordAnalysis.trim(table[0]);
        joinResult.as = SelectKeywordAnalysis.trim(table[1]) || '';
        joinResult.cond = SelectKeywordAnalysis.trim(strParts[1]);

        return joinResult;
    }

    static WHERE(str) {
        return CondParser.parse(str);
    }

    static ORDER_BY(str) {
        return SelectKeywordAnalysis.SELECT(str, true);
    }

    static GROUP_BY(str) {
        return SelectKeywordAnalysis.SELECT(str);
    }

    static PIVOT(str) {
        const strParts = str.split(',');
        const pivotResult = [];

        strParts.forEach((item, _key) => {
            const pivotOn = /([\w.]+)/gi;
            const pivotData = pivotOn.exec(item);
            if (pivotData !== null) {
                const tmp = {};
                tmp.name = SelectKeywordAnalysis.trim(pivotData[1]);
                tmp.as = "";
                pivotResult.push(tmp);
            }
        });

        return pivotResult;
    }

    static LIMIT(str) {
        const limitResult = {};
        limitResult.nb = Number(str);
        limitResult.from = 0;
        return limitResult;
    }

    static HAVING(str) {
        return CondParser.parse(str);
    }

    static UNION(str) {
        return SelectKeywordAnalysis.trim(str);
    }

    static UNION_ALL(str) {
        return SelectKeywordAnalysis.trim(str);
    }

    static INTERSECT(str) {
        return SelectKeywordAnalysis.trim(str);
    }

    static EXCEPT(str) {
        return SelectKeywordAnalysis.trim(str);
    }

    /**
     * 
     * @param {String} selectField 
     * @returns {Object}
     */
    static parseForCorrelatedSubQuery(selectField) {
        let subQueryAst = null;

        const regExp = /\(\s*(SELECT[\s\S]+)\)/;
        const matches = regExp.exec(selectField.toUpperCase());

        if (matches !== null && matches.length > 1) {
            subQueryAst = SqlParse.sql2ast(matches[1]);
        }

        return subQueryAst;
    }

    // Split a string using a separator, only if this separator isn't beetween brackets
    /**
     * 
     * @param {String} separator 
     * @param {String} str 
     * @returns {String[]}
     */
    static protect_split(separator, str) {
        const sep = '######';

        let inQuotedString = false;
        let quoteChar = "";
        let bracketCount = 0;
        let newStr = "";
        for (const c of str) {
            if (!inQuotedString && /['"`]/.test(c)) {
                inQuotedString = true;
                quoteChar = c;
            }
            else if (inQuotedString && c === quoteChar) {
                inQuotedString = false;
            }
            else if (!inQuotedString && c === '(') {
                bracketCount++;
            }
            else if (!inQuotedString && c === ')') {
                bracketCount--;
            }

            if (c === separator && (bracketCount > 0 || inQuotedString)) {
                newStr += sep;
            }
            else {
                newStr += c;
            }
        }

        let strParts = newStr.split(separator);
        strParts = strParts.map(item => SelectKeywordAnalysis.trim(item.replace(new RegExp(sep, 'g'), separator)));

        return strParts;
    }

    static trim(str) {
        if (typeof str === 'string')
            return str.trim();
        return str;
    }

    /**
    * If an ALIAS is specified after 'AS', return the field/table name and the alias.
    * @param {String} item 
    * @returns {String[]}
    */
    static getNameAndAlias(item) {
        let realName = item;
        let alias = "";
        const lastAs = SelectKeywordAnalysis.lastIndexOfOutsideLiteral(item.toUpperCase(), " AS ");
        if (lastAs !== -1) {
            const subStr = item.substring(lastAs + 4).trim();
            if (subStr.length > 0) {
                alias = subStr;
                //  Remove quotes, if any.
                if ((subStr.startsWith("'") && subStr.endsWith("'")) ||
                    (subStr.startsWith('"') && subStr.endsWith('"')) ||
                    (subStr.startsWith('[') && subStr.endsWith(']')))
                    alias = subStr.substring(1, subStr.length - 1);

                //  Remove everything after 'AS'.
                realName = item.substring(0, lastAs);
            }
        }

        return [realName, alias];
    }

    /**
     * 
     * @param {String} srcString 
     * @param {String} searchString 
     * @returns {Number}
     */
    static lastIndexOfOutsideLiteral(srcString, searchString) {
        let index = -1;
        let inQuote = "";

        for (let i = 0; i < srcString.length; i++) {
            const ch = srcString.charAt(i);

            if (inQuote !== "") {
                //  The ending quote.
                if ((inQuote === "'" && ch === "'") || (inQuote === '"' && ch === '"') || (inQuote === "[" && ch === "]"))
                    inQuote = "";
            }
            else if ("\"'[".indexOf(ch) !== -1) {
                //  The starting quote.
                inQuote = ch;
            }
            else if (srcString.substring(i).startsWith(searchString)) {
                //  Matched search.
                index = i;
            }
        }

        return index;
    }
}


/** 
 * Interface for loading table data either from CACHE or SHEET. 
 * @class
 * @classdesc
 * * Automatically load table data from a **CACHE** or **SHEET** <br>
 * * In all cases, if the cache has expired, the data is read from the sheet. 
 * <br>
 * 
 * | Cache Seconds | Description |
 * | ---           | ---         |
 * | 0             | Data is not cached and always read directly from SHEET |
 * | <= 21600      | Data read from SHEETS cache if it has not expired |
 * | > 21600       | Data read from Google Sheets Script Settings |
 * 
 */
class TableData {       //  skipcq: JS-0128
    /**
    * Retrieve table data from SHEET or CACHE.
    * @param {String} namedRange - Location of table data.  Either a) SHEET Name, b) Named Range, c) A1 sheet notation.
    * @param {Number} cacheSeconds - 0s Reads directly from sheet. > 21600s Sets in SCRIPT settings, else CacheService 
    * @returns {any[][]}
    */
    static loadTableData(namedRange, cacheSeconds = 0) {
        if (typeof namedRange === 'undefined' || namedRange === "")
            return [];

        Logger.log(`loadTableData: ${namedRange}. Seconds=${cacheSeconds}`);

        let tempData = Table.removeEmptyRecordsAtEndOfTable(TableData.getValuesCached(namedRange, cacheSeconds));

        return tempData;
    }

    /**
     * Reads a RANGE of values.
     * @param {String} namedRange 
     * @param {Number} seconds 
     * @returns {any[][]}
     */
    static getValuesCached(namedRange, seconds) {
        let cache = {};
        let cacheSeconds = seconds;

        if (cacheSeconds <= 0) {
            return TableData.loadValuesFromRangeOrSheet(namedRange);
        }
        else if (cacheSeconds > 21600) {
            cache = new ScriptSettings();
            if (TableData.isTimeToRunLongCacheExpiry()) {
                cache.expire(false);
                TableData.setLongCacheExpiry();
            }
            cacheSeconds = cacheSeconds / 86400;  //  ScriptSettings put() wants days to hold.
        }
        else {
            cache = CacheService.getScriptCache();
        }

        let arrData = TableData.cacheGetArray(cache, namedRange);
        if (arrData !== null) {
            Logger.log(`Found in CACHE: ${namedRange}. Items=${arrData.length}`);
            return arrData;
        }

        Logger.log(`Not in cache: ${namedRange}`);

        arrData = TableData.lockLoadAndCache(cache, namedRange, cacheSeconds);

        return arrData;
    }

    /**
     * Is it time to run the long term cache expiry check?
     * @returns {Boolean}
     */
    static isTimeToRunLongCacheExpiry() {
        const shortCache = CacheService.getScriptCache();
        return shortCache.get("LONG_CACHE_EXPIRY") === null;
    }

    /**
     * The long term expiry check is done every 21,000 seconds.  Set the clock now!
     */
    static setLongCacheExpiry() {
        const shortCache = CacheService.getScriptCache();
        shortCache.put("LONG_CACHE_EXPIRY", 'true', 21000);
    }

    /**
     * In the interest of testing, force the expiry check.
     * It does not mean items in cache will be removed - just 
     * forces a check.
     */
    static forceLongCacheExpiryCheck() {
        const shortCache = CacheService.getScriptCache();
        if (shortCache.get("LONG_CACHE_EXPIRY") !== null) {
            shortCache.remove("LONG_CACHE_EXPIRY");
        }
    }

    /**
     * Reads a single cell.
     * @param {String} namedRange 
     * @param {Number} seconds 
     * @returns {any}
     */
    static getValueCached(namedRange, seconds = 60) {
        const cache = CacheService.getScriptCache();

        let singleData = cache.get(namedRange);

        if (singleData === null) {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            singleData = ss.getRangeByName(namedRange).getValue();
            cache.put(namedRange, JSON.stringify(singleData), seconds);
        }
        else {
            singleData = JSON.parse(singleData);
            const tempArr = [[singleData]];
            TableData.fixJSONdates(tempArr);
            singleData = tempArr[0][0];
        }

        return singleData;
    }

    /**
     * For updating a sheet VALUE that may be later read from cache.
     * @param {String} namedRange 
     * @param {any} singleData 
     * @param {Number} seconds 
     */
    static setValueCached(namedRange, singleData, seconds = 60) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        ss.getRangeByName(namedRange).setValue(singleData);
        let cache = null;

        if (seconds === 0) {
            return;
        }
        else if (seconds > 21600) {
            cache = new ScriptSettings();
        }
        else {
            cache = CacheService.getScriptCache();
        }
        cache.put(namedRange, JSON.stringify(singleData), seconds);
    }

    /**
     * 
     * @param {String} namedRange 
     * @param {any[][]} arrData 
     * @param {Number} seconds 
     */
    static setValuesCached(namedRange, arrData, seconds = 60) {
        const cache = CacheService.getScriptCache();

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        ss.getRangeByName(namedRange).setValues(arrData);
        cache.put(namedRange, JSON.stringify(arrData), seconds)
    }

    /**
     * Check if data from cache is in error.
     * @param {any[][]} arrData 
     * @returns {Boolean}
     */
    static verifyCachedData(arrData) {
        let verified = true;

        for (const rowData of arrData) {
            for (const fieldData of rowData) {
                if (fieldData === "#ERROR!") {
                    Logger.log("Reading from CACHE has found '#ERROR!'.  Re-Loading...");
                    verified = false;
                    break;
                }
            }
        }

        return verified;
    }

    /**
     * Checks if this range is loading elsewhere (i.e. from another call to custom function)
     * @param {String} namedRange
     * @returns {Boolean} 
     */
    static isRangeLoading(cache, namedRange) {
        let loading = false;
        const cacheData = cache.get(TableData.cacheStatusName(namedRange));

        if (cacheData !== null && cacheData === TABLE.LOADING) {
            loading = true;
        }

        Logger.log(`isRangeLoading: ${namedRange}. Status: ${loading}`);

        return loading;
    }

    /**
     * Retrieve data from cache after it has loaded elsewhere.
     * @param {Object} cache 
     * @param {String} namedRange 
     * @param {Number} cacheSeconds - How long to cache results.
     * @returns {any[][]}
     */
    static waitForRangeToLoad(cache, namedRange, cacheSeconds) {
        const start = new Date().getTime();
        let current = new Date().getTime();

        Logger.log(`waitForRangeToLoad() - Start: ${namedRange}`);
        while (TableData.isRangeLoading(cache, namedRange) && (current - start) < 10000) {
            Utilities.sleep(250);
            current = new Date().getTime();
        }
        Logger.log("waitForRangeToLoad() - End");

        let arrData = TableData.cacheGetArray(cache, namedRange);

        //  Give up and load from SHEETS directly.
        if (arrData === null) {
            Logger.log(`waitForRangeToLoad - give up.  Read directly. ${namedRange}`);
            arrData = TableData.loadValuesFromRangeOrSheet(namedRange);

            if (TableData.isRangeLoading(cache, namedRange)) {
                //  Other process probably timed out and left status hanging.
                TableData.cachePutArray(cache, namedRange, cacheSeconds, arrData);
            }
        }

        return arrData;
    }

    /**
     * Read range of value from sheet and cache.
     * @param {Object} cache - cache object can vary depending where the data is stored.
     * @param {String} namedRange 
     * @param {Number} cacheSeconds 
     * @returns {any[][]} - data from range
     */
    static lockLoadAndCache(cache, namedRange, cacheSeconds) {
        //  Only change our CACHE STATUS if we have a lock.
        const lock = LockService.getScriptLock();
        try {
            lock.waitLock(10000); // wait 10 seconds for others' use of the code section and lock to stop and then proceed
        } catch (e) {
            throw new Error("Cache lock failed");
        }

        //  It is possible that just before getting the lock, another process started caching.
        if (TableData.isRangeLoading(cache, namedRange)) {
            lock.releaseLock();
            return TableData.waitForRangeToLoad(cache, namedRange, cacheSeconds);
        }

        //  Mark the status for this named range that loading is in progress.
        cache.put(TableData.cacheStatusName(namedRange), TABLE.LOADING, 15);
        lock.releaseLock();

        //  Load data from SHEETS.
        const arrData = TableData.loadValuesFromRangeOrSheet(namedRange);

        Logger.log(`Just LOADED from SHEET: ${arrData.length}`);

        TableData.cachePutArray(cache, namedRange, cacheSeconds, arrData);

        return arrData;
    }

    /**
     * Read sheet data into double array.
     * @param {String} namedRange - named range, A1 notation or sheet name
     * @returns {any[][]} - table data.
     */
    static loadValuesFromRangeOrSheet(namedRange) {
        let tableNamedRange = namedRange;
        let output = [];

        try {
            const sheetNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(tableNamedRange);

            if (sheetNamedRange === null) {
                //  This may be a SHEET NAME, so try getting SHEET RANGE.
                if (tableNamedRange.startsWith("'") && tableNamedRange.endsWith("'")) {
                    tableNamedRange = tableNamedRange.substring(1, tableNamedRange.length - 1);
                }
                let sheetHandle = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tableNamedRange);

                //  Actual sheet may have spaces in name.  The SQL must reference that table with
                //  underscores replacing those spaces.
                if (sheetHandle === null && tableNamedRange.indexOf("_") !== -1) {
                    tableNamedRange = tableNamedRange.replace(/_/g, " ");
                    sheetHandle = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tableNamedRange);
                }

                if (sheetHandle === null) {
                    throw new Error(`Invalid table range specified:  ${tableNamedRange}`);
                }

                const lastColumn = sheetHandle.getLastColumn();
                const lastRow = sheetHandle.getLastRow();
                output = sheetHandle.getSheetValues(1, 1, lastRow, lastColumn);
            }
            else {
                // @ts-ignore
                output = sheetNamedRange.getValues();
            }
        }
        catch (ex) {
            throw new Error(`Error reading table data: ${tableNamedRange}`);
        }

        return output;
    }

    /**
     * Takes array data to be cached, breaks up into chunks if necessary, puts each chunk into cache and updates status.
     * @param {Object} cache 
     * @param {String} namedRange 
     * @param {Number} cacheSeconds 
     * @param {any[][]} arrData 
     */
    static cachePutArray(cache, namedRange, cacheSeconds, arrData) {
        const cacheStatusName = TableData.cacheStatusName(namedRange);
        const json = JSON.stringify(arrData);

        //  Split up data (for re-assembly on get() later)
        let splitCount = (json.length / (100 * 1024)) * 1.3;    // 1.3 - assumes some blocks may be bigger.
        splitCount = splitCount < 1 ? 1 : splitCount;
        const arrayLength = Math.ceil(arrData.length / splitCount);
        const putObject = {};
        let blockCount = 0;
        let startIndex = 0;
        while (startIndex < arrData.length) {
            const arrayBlock = arrData.slice(startIndex, startIndex + arrayLength);
            blockCount++;
            startIndex += arrayLength;
            putObject[`${namedRange}:${blockCount.toString()}`] = JSON.stringify(arrayBlock);
        }

        //  Update status that cache is updated.
        const lock = LockService.getScriptLock();
        try {
            lock.waitLock(10000); // wait 10 seconds for others' use of the code section and lock to stop and then proceed
        } catch (e) {
            throw new Error("Cache lock failed");
        }
        cache.putAll(putObject, cacheSeconds);
        cache.put(cacheStatusName, TABLE.BLOCKS + blockCount.toString(), cacheSeconds);

        Logger.log(`Writing STATUS: ${cacheStatusName}. Value=${TABLE.BLOCKS}${blockCount.toString()}. seconds=${cacheSeconds}. Items=${arrData.length}`);

        lock.releaseLock();
    }

    /**
     * Reads cache for range, and re-assembles blocks into return array of data.
     * @param {Object} cache 
     * @param {String} namedRange 
     * @returns {any[][]}
     */
    static cacheGetArray(cache, namedRange) {
        let arrData = [];

        const cacheStatusName = TableData.cacheStatusName(namedRange);
        const cacheStatus = cache.get(cacheStatusName);
        if (cacheStatus === null) {
            Logger.log(`Named Range Cache Status not found = ${cacheStatusName}`);
            return null;
        }

        Logger.log(`Cache Status: ${cacheStatusName}. Value=${cacheStatus}`);
        if (cacheStatus === TABLE.LOADING) {
            return null;
        }

        const blockStr = cacheStatus.substring(cacheStatus.indexOf(TABLE.BLOCKS) + TABLE.BLOCKS.length);
        if (blockStr !== "") {
            const blocks = Number(blockStr);
            for (let i = 1; i <= blocks; i++) {
                const blockName = `${namedRange}:${i.toString()}`;
                const jsonData = cache.get(blockName);

                if (jsonData === null) {
                    Logger.log(`Named Range Part not found. R=${blockName}`);
                    return null;
                }

                const partArr = JSON.parse(jsonData);
                if (TableData.verifyCachedData(partArr)) {
                    arrData = arrData.concat(partArr);
                }
                else {
                    Logger.log(`Failed to verify named range: ${blockName}`);
                    return null;
                }
            }

        }
        Logger.log(`Just LOADED From CACHE: ${namedRange}. Items=${arrData.length}`);

        //  The conversion to JSON causes SHEET DATES to be converted to a string.
        //  This converts any DATE STRINGS back to javascript date.
        TableData.fixJSONdates(arrData);

        return arrData;
    }

    /**
     * 
     * @param {any[][]} arrData 
     */
    static fixJSONdates(arrData) {
        const ISO_8601_FULL = /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(\.\d+)?(([+-]\d\d:\d\d)|Z)?$/i

        for (const row of arrData) {
            for (let i = 0; i < row.length; i++) {
                const testStr = row[i];
                if (ISO_8601_FULL.test(testStr)) {
                    row[i] = new Date(testStr);
                }
            }
        }
    }

    /**
     * 
     * @param {String} namedRange 
     * @returns {String}
     */
    static cacheStatusName(namedRange) {
        return namedRange + TABLE.STATUS;
    }
}

const TABLE = {
    STATUS: "__STATUS__",
    LOADING: "LOADING",
    BLOCKS: "BLOCKS="
}



/** Stores settings for the SCRIPT.  Long term cache storage for small tables.  */
class ScriptSettings {      //  skipcq: JS-0128
    /**
     * For storing cache data for very long periods of time.
     */
    constructor() {
        this.scriptProperties = PropertiesService.getScriptProperties();
    }

    /**
     * Get script property using key.  If not found, returns null.
     * @param {String} propertyKey 
     * @returns {any}
     */
    get(propertyKey) {
        const myData = this.scriptProperties.getProperty(propertyKey);

        if (myData === null)
            return null;

        /** @type {PropertyData} */
        const myPropertyData = JSON.parse(myData);

        return PropertyData.getData(myPropertyData);
    }

    /**
     * Put data into our PROPERTY cache, which can be held for long periods of time.
     * @param {String} propertyKey - key to finding property data.
     * @param {any} propertyData - value.  Any object can be saved..
     * @param {Number} daysToHold - number of days to hold before item is expired.
     */
    put(propertyKey, propertyData, daysToHold = 1) {
        //  Create our object with an expiry time.
        const objData = new PropertyData(propertyData, daysToHold);

        //  Our property needs to be a string
        const jsonData = JSON.stringify(objData);

        try {
            this.scriptProperties.setProperty(propertyKey, jsonData);
        }
        catch (ex) {
            throw new Error("Cache Limit Exceeded.  Long cache times have limited storage available.  Only cache small tables for long periods.");
        }
    }

    /**
     * 
     * @param {Object} propertyDataObject 
     * @param {Number} daysToHold 
     */
    putAll(propertyDataObject, daysToHold = 1) {
        const keys = Object.keys(propertyDataObject);

        for (const key of keys) {
            this.put(key, propertyDataObject[key], daysToHold);
        }
    }

    /**
     * Removes script settings that have expired.
     * @param {Boolean} deleteAll - true - removes ALL script settings regardless of expiry time.
     */
    expire(deleteAll) {
        const allKeys = this.scriptProperties.getKeys();

        for (const key of allKeys) {
            const myData = this.scriptProperties.getProperty(key);

            if (myData !== null) {
                let propertyValue = null;
                try {
                    propertyValue = JSON.parse(myData);
                }
                catch (e) {
                    Logger.log(`Script property data is not JSON. key=${key}`);
                }

                if (propertyValue !== null && (PropertyData.isExpired(propertyValue) || deleteAll)) {
                    this.scriptProperties.deleteProperty(key);
                    Logger.log(`Removing expired SCRIPT PROPERTY: key=${key}`);
                }
            }
        }
    }
}

/** Converts data into JSON for getting/setting in ScriptSettings. */
class PropertyData {
    /**
     * 
     * @param {any} propertyData 
     * @param {Number} daysToHold 
     */
    constructor(propertyData, daysToHold) {
        const someDate = new Date();

        /** @property {String} */
        this.myData = JSON.stringify(propertyData);
        /** @property {Date} */
        this.expiry = someDate.setMinutes(someDate.getMinutes() + daysToHold * 1440);
    }

    /**
     * 
     * @param {PropertyData} obj 
     * @returns {any}
     */
    static getData(obj) {
        let value = null;
        try {
            if (!PropertyData.isExpired(obj))
                value = JSON.parse(obj.myData);
        }
        catch (ex) {
            Logger.log(`Invalid property value.  Not JSON: ${ex.toString()}`);
        }

        return value;
    }

    /**
     * 
     * @param {PropertyData} obj 
     * @returns {Boolean}
     */
    static isExpired(obj) {
        const someDate = new Date();
        const expiryDate = new Date(obj.expiry);
        return (expiryDate.getTime() < someDate.getTime())
    }
}

