// @author Sourabh Choraria https://script.gs/google-sheets-custom-functions-for-wannabe-domain-investors/
/**
 * Returns data about an available .com domain.
 *
 * @param {string} name The name of the domain.
 * @return {Array} Registrar name, registration & expiration date of a .com domain.
 * @customfunction
 */
function DOT_COM_DATA(name) {
  const nameComponents = name.replace(/\s+/g, '').split(".");
  if (nameComponents.length > 2) return "INVALID INPUT";
  if (nameComponents.length == 2 && nameComponents[1] != "com") return "TLD NOT SUPPORTED";
  name = nameComponents[0];
  const url = `https://rdap.verisign.com/com/v1/domain/${name}.com`;
  const response = UrlFetchApp.fetch(url,{ muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return "AVAILABLE";
  let comData = [];
  const jsonData = JSON.parse(response.getContentText());
  const registrar = jsonData.entities[0].vcardArray[1][1][3];
  const registrationDate = jsonData.events[0].eventDate.replace("T"," ").replace("Z","");
  const expirationDate = jsonData.events[1].eventDate.replace("T"," ").replace("Z","");
  comData.push([registrar, registrationDate, expirationDate]);
  return comData;
}