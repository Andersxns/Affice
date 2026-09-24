/** Loads every function module so they register themselves. */
import './math';
import './stats';
import './text';
import './logical';
import './lookup';
import './datetime';
import './info';
import './financial';
import './engineering';
import './distributions';
import './database';

export { allFunctions, getFunction, hasFunction, type Category, type FnDef } from './registry';
