// @author Amit Agarwal https://www.labnol.org/google-maps-sheets-200817
/**
 * Calculate the travel time between two locations
 * on Google Maps.
 *
 * =GOOGLEMAPS_DURATION("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The time in minutes
 * @customFunction
 */
const GOOGLEMAPS_DURATION = (origin, destination, mode = 'driving') => {
  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
  if (!data) {
    throw new Error('No route found!');
  }
  const { legs: [{ duration: { text: time } } = {}] = [] } = data;
  return time;
};