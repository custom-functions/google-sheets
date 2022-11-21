/**
 * Calculate the distance between two locations on Google Maps.
 *
 * =GOOGLEMAPS_DISTANCE("NY 10005", "Hoboken NJ", "walking")
 *
 * @author Amit Agarwal
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The distance in miles
 * @customfunction
 */
const GOOGLEMAPS_DISTANCE = (origin, destination, mode) => {
  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();

  if (!data) {
    throw new Error('No route found!');
  }

  const { legs: [{ distance: { text: distance } } = {}] = [] } = data;
  return distance;
};
