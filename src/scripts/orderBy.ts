/*--------------------------------------------------------------------------
   Definitions of the supported orderBy parameters.
 -------------------------------------------------------------------------*/

export enum Parameter {
	lastModifiedTime = "lastModifiedTime",
	createdTime = "createdTime",
	name = "name"
}

export enum Direction {
	asc = "asc",
	desc = "desc"
}

export class OrderBy {
	public parameter: Parameter;
	public direction: Direction;
}
