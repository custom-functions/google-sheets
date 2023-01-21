// @author Nwachukwu Ujubuonu https://github.com/Hemephelus/google-sheets-custom-function
/**
 * Returns unique rows in the provided source range and selected columns, discarding duplicates. Rows are returned in the order in which they first appear in the source range. 
 * 
 * @param {C1:D2} range The data to filter by unique entries.
 * @param {"1,2"} cols The columns you want to filter by.
 *
 * @return Unique rows in the provided source range and selected columns.
 * @customfunction
 */ 
function SUPER_UNIQUE(range,cols) {
  cols = cols.replace(" ","").split(',')
  let newRange = {}

  for(let i = 0; i < range.length; i++){
    let subRange = ''
    for(let j = 0; j < cols.length; j++){
      subRange = subRange+range[i][cols[j]-1]
    }
     newRange[subRange] = range[i]

  }

 return Object.values(newRange)
  
  
}
