// @author Amit Agarwal https://www.labnol.org/google-maps-sheets-200817
/**
 * Use Reverse Geocoding to get the address of
 * a point location (latitude, longitude) on Google Maps.
 *
 * =GOOGLEMAPS_REVERSEGEOCODE(latitude, longitude)
 *
 * @param {String} latitude The latitude to lookup.
 * @param {String} longitude The longitude to lookup.
 * @return {String} The postal address of the point.
 * @customFunction
 */
const GOOGLEMAPS_REVERSEGEOCODE = (latitude, longitude) => {
  const { results: [data = {}] = [] } = Maps.newGeocoder().reverseGeocode(latitude, longitude);
  return data.formatted_address;
};