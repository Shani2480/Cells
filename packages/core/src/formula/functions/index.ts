// Registers every function. Importing this module (the formula index does) fills FUNCTIONS.
import "./math"
import "./stats"
import "./logical"
import "./info"
import "./text"
import "./datetime"
import "./lookup"
import "./financial"
import "./aggregate"
import "./forecast"

export { FUNCTIONS, isVolatile } from "./registry"
export { compileCriteria, wildcardPattern } from "./criteria"
