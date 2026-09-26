import pkg from '../../package.json';

declare const __COMMIT__: string;
declare const __BUILD_ZEIT__: string;

export const VERSION: string = pkg.version;

/** Vollständiger Commit-SHA des Builds (leer, wenn unbekannt; Endung „-geändert" bei lokalen Änderungen). */
export const COMMIT: string = __COMMIT__;

/** Kurzform des Commits, wie GitHub sie anzeigt. */
export const COMMIT_KURZ: string = COMMIT.replace(/^([0-9a-f]{7})[0-9a-f]*/, '$1');

/** Zeitpunkt des Builds (ISO 8601). */
export const BUILD_ZEIT: string = __BUILD_ZEIT__;
