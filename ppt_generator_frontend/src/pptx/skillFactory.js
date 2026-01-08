export const skillFactoryName = "Skill Factory";

/**
 * Generates an array of 4 slides for the given skill factory name.
 * These are placeholder slides representing the factory section.
 * @param {string} factoryLabel e.g. "Skill Factory 1"
 */
export function generateSkillFactorySlides(factoryLabel) {
  // Return 4 dummy slides with unique titles
  return [
    { type: "factory", factory: factoryLabel, title: `${factoryLabel} - Slide 1` },
    { type: "factory", factory: factoryLabel, title: `${factoryLabel} - Slide 2` },
    { type: "factory", factory: factoryLabel, title: `${factoryLabel} - Slide 3` },
    { type: "factory", factory: factoryLabel, title: `${factoryLabel} - Slide 4` },
  ];
}
