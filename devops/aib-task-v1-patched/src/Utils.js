"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
class Utils {
    static IsEqual(a, b) {
        return a.toLowerCase() == b.toLowerCase();
    }
    // v1-patched: needed for progress polling delay
    static sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}
exports.default = Utils;
