"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const tl = require("azure-pipelines-task-lib/task");
const ImageBuilder_1 = __importDefault(require("./ImageBuilder"));
async function run() {
    var ib = new ImageBuilder_1.default();
    await ib.execute();
}
run().then()
    .catch((error) => tl.setResult(tl.TaskResult.Failed, error));
