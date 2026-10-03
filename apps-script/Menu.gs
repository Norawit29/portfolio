/**
 * Entry points. If the project already has an onOpen() or doGet(), delete
 * these two and call pfOnOpen(e) / return pfDoGet(e) from the existing ones.
 */
function onOpen(e) {
  pfOnOpen(e);
}

function doGet(e) {
  return pfDoGet(e);
}
