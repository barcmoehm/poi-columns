Used for working with Apache POI. Transforms the standard representation (Row) to a column representation of the data.

Standard Structure POI:  [{A,B,C},{1,2,3},{1,2,3},{1,2,3}]
Structure of POI columns: [A:{1,1,1},B:{2,2,2},C:{3,3,3}]

This allows for filtering and other operations. Most likely, there is a better solution somewhere on the internet. Don't use this with big spreadsheets.
