// @author Amit Agarwal https://www.labnol.org/google-maps-sheets-200817
/**
 * Get the latitude and longitude of any
 * address on Google Maps.
 *
 * =GOOGLEMAPS_LATLONG("10 Hanover Square, NY")
 *
 * @param {String} address The address to lookup.
 * @return {String} The latitude and longitude of the address.
 * @customFunction
 */
const GOOGLEMAPS_LATLONG = (address) => {
  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) {
    throw new Error('Address not found!');
  }
  const { geometry: { location: { lat, lng } } = {} } = data;
  return `${lat}, ${lng}`;
};