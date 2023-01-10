# Permeator

Migration of the previous VB6 extension to VB.NET and refinement of the permeation function.

The extension first calculates atomic diffusion flows for H, D and T following mainly the mixed formula for Ackerman1972 which calculates the contribution of the different species (e.g., HD, HT, and H2 for the H diffusion) in the X and Y coefficients (currently we assume outlet pressure = 0, i.e., Y = 0) using partial pressures.

Before, the permeation of components was hard coded and only for diatomic hydrogen and tritium. Now it can handle the six isotopologues in a more refined way. For this, it uses heuristics, depending on the initial composition. If we have an excess of every permeating molecule, the extension distributes them proportionally. If it is only some of them, we distribute it starting from the lower flow rate so that we first exhaust the lower ones. 

To make it simpler (and especially for the mass balance of HYSYS, but this is still to be tested), we don’t make changes in molecules, i.e., if the inlet has only HD, we make outlet ONLY HD and not H2 or D2. However, these assumptions and the heuristics are not perfect, and we might not permeate correctly all the initially calculated flow,

*	if one of the molecular streams is lower and exhausts in the middle of the permeation. This effect might be weaker with the discretization of the permeator.
*	if we have only a bimolecular species, i.e., HD. Since H and D have different permeabilities. This case should be avoided.

In terms of performing flash, we removed flash during Execute since it was not entirely necessary. It was sometimes problematic because the fluid objects sometimes struggle and are not “UpToDate”, and the aspentech example permeator does not use it anyway. For Dynamics, it remains the question of using flash or not for the Energy and Composition steps and if it makes any difference for the extra computational cost.
