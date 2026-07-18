"use strict";

export default class Utils
{
    public static IsEqual(a: string, b: string): boolean
    {
        return a.toLowerCase() == b.toLowerCase()
    }

    // v1-patched: needed for progress polling delay
    public static sleep(ms: number): Promise<void>
    {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}
