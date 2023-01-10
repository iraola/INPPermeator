# Issues

* Execute(). To flash or not to flash? Flash is sometimes problematic, and not only that but it clutters the code when it might not always be necessary. Taking the example of the aspentech permeator extension, you can follow the next procedure:
    1. Calculate permeation flows (separately in Double() arrays)
    2. Assign to the edf streams through the `Calculate` method: e.g. `edfRetentate.ComponentMolarFlow.Calculate(<put here the array of component flow>)`, etc.
    3. Assign also other especifications that would be needed for a flash (e.g. temperature-pressure, pressure-enthalpy, ...) again via the `.Calculate` methods.
    4. Do `Balance()`
    5. Perform fluids `UpToDate` check and do `SolveComplete()`

    This way `Balance` kind of performs the flash if the strams are correctly specified.

    **We will see how this solving procedure without explicit flash works for Dynamic simulation.** In the past I used fluids and flash only for the Energy and Composition steps, which I guess it makes sense; but it might not even be necessary, let's check that.
* Type of flash. Definitely we prefer to do a **TP flash**, even though we lose some accuracy of the "valve" behavior. But handling two output streams and with very low pressures on the permeate side, sometimes a PH is not even able to solve.
* Execute(). Issues with the fluid readiness check before doing `myContainer.SolveComplete()`:

    `If edfInlet.DuplicateFluid.IsUpToDate And edfPermeate.DuplicateFluid.IsUpToDate And edfRetentate.DuplicateFluid.IsUpToDate Then`
 
    It is good practice to use it, but need to be sure the streams and correctly defined and solved after doing `myContainer.Balance()`
* Heuristics for distributing atomic diffusing flow into molecular gas flow. Issues:
  - one of the molecular flows is lower and exhausts in the middle of the permeation. This effect might be weaker with the discretization of the permeator.
  - only heteronuclear flow + with different permeabilities between atoms might cause the lower permeability species to limit permeation.