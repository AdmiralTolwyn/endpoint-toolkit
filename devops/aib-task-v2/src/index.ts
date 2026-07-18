import tl = require('azure-pipelines-task-lib/task');
import ImageBuilder from './ImageBuilder';

async function run() : Promise<void>{
     var ib = new ImageBuilder();
     await ib.execute();
}

run().then()
     .catch((error) => tl.setResult(tl.TaskResult.Failed, error));
