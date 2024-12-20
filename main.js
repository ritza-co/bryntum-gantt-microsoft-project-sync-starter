import { Gantt } from '@bryntum/gantt';
import '@bryntum/gantt/gantt.stockholm.css';
import { signIn } from './auth.js';
import { abandonOperationSet, createOperationSet, createProjectTask, createProjectTaskDependency, deleteProjectTask, deleteProjectTaskDependency, executeOperationSet, getProjectTaskDependencies, getProjectTasks, updateProjectTask, waitForTaskCreationThenGetId } from './crudFunctions.js';
import CustomTaskModel from './lib/CustomTaskModel.js';

const signInLink = document.getElementById('signin');

const gantt = new Gantt({
    appendTo   : 'gantt',
    viewPreset : 'weekAndMonth',
    timeZone   : 'UTC',
    date       : new Date(2024, 10, 1),
    project    : {
        taskStore : {
            transformFlatData : true,
            modelClass        : CustomTaskModel
        },
        writeAllFields : true
    },
    listeners : {
        dataChange : function(event) {
            updateMicrosoftProject(event);
        }
    }
});
async function displayUI() {
    await signIn();

    // Hide sign in link and initial UI
    signInLink.style = 'display: none';
    const content = document.getElementById('content');
    content.style = 'display: block';

    try {
        const projectTasksPromise = getProjectTasks();
        const getProjectDependenciesPromise = getProjectTaskDependencies();
        const [projectTasks, projectDependencies] = await Promise.all([
            projectTasksPromise,
            getProjectDependenciesPromise
        ]);

        const ganttTasks = [];
        projectTasks.forEach((event) => {
            const startDateUTC = new Date(event.msdyn_start);
            // Convert to local timezone
            const startDateLocal = new Date(
                startDateUTC.getTime() - startDateUTC.getTimezoneOffset() * 60000
            );
            const finishDateUTC = new Date(event.msdyn_finish);
            // Convert to local timezone
            const finishDateLocal = new Date(
                finishDateUTC.getTime() - finishDateUTC.getTimezoneOffset() * 60000
            );
            ganttTasks.push({
                id                    : event.msdyn_projecttaskid,
                parentId              : event._msdyn_parenttask_value,
                name                  : event.msdyn_subject,
                startDate             : startDateLocal,
                endDate               : finishDateLocal,
                percentDone           : event.msdyn_progress * 100,
                msdyn_displaysequence : parseInt(event.msdyn_displaysequence),
                manuallyScheduled     : true,
                outlineLevel          : event.msdyn_outlinelevel
            });
        });
        ganttTasks.sort((a, b) => a.msdyn_displaysequence - b.msdyn_displaysequence);
        gantt.project.tasks = ganttTasks;
        const ganttDependencies = [];
        projectDependencies.forEach((dep) => {
            ganttDependencies.push({
                id   : dep.msdyn_projecttaskdependencyid,
                from : dep._msdyn_predecessortask_value,
                to   : dep._msdyn_successortask_value
            });
        });
        gantt.project.dependencies = ganttDependencies;
    }
    catch (error) {
        console.error('Error:', error);
    }
}

signInLink.addEventListener('click', displayUI);

async function updateMicrosoftProject({ action, record, store, records }) {
    const storeId = store.id;
    if (storeId === 'tasks') {
        if (action === 'update') {
            if (`${record.id}`.startsWith('_generated')) {
                if (!record.name) return;
                let operationSetId = '';
                const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                const projectBucketId = import.meta.env.VITE_MSDYN_PROJECTBUCKET_VALUE;
                const description = 'Create operation set for new project task';
                try {
                    gantt.maskBody('Creating task...');
                    operationSetId = await createOperationSet(projectId, description);
                    let msdyn_displaysequence = null;
                    if (!record.previousSibling) {
                        msdyn_displaysequence = record.nextSibling.msdyn_displaysequence / 2;
                    }
                    if (!record.nextSibling) {
                        msdyn_displaysequence = record.previousSibling.msdyn_displaysequence + 1;
                    }
                    if (record.previousSibling && record.nextSibling) {
                        msdyn_displaysequence = (record.previousSibling.msdyn_displaysequence + record.nextSibling.msdyn_displaysequence) / 2;
                    }
                    // Add a small fraction to avoid value conflicts
                    msdyn_displaysequence += 0.01;
                    if (msdyn_displaysequence <= 1) {
                        msdyn_displaysequence = 1.1;
                    }
                    // round to maximum 9 decimal places
                    msdyn_displaysequence = Number(msdyn_displaysequence.toFixed(9));
                    const createProjectTaskResponse = await createProjectTask(projectId, projectBucketId, operationSetId, record, msdyn_displaysequence);
                    const newId = JSON.parse(createProjectTaskResponse.OperationSetResponse)['<OperationSetResponses>k__BackingField'][3].Value;

                    await executeOperationSet(operationSetId);
                    // wait 500 ms
                    // fetch the newly created task by id to get its new display sequence - it is not returned by createProjectTask
                    const new_msdyn_displaysequence = await waitForTaskCreationThenGetId(newId);
                    // update id
                    gantt.project.taskStore.applyChangeset({
                        updated : [
                            // Will set proper id for added task
                            {
                                $PhantomId            : record.id,
                                id                    : newId,
                                msdyn_displaysequence : new_msdyn_displaysequence
                            }
                        ]
                    });
                    return;
                }
                catch (error) {
                    await abandonOperationSet(operationSetId);
                    console.error('Error:', error);
                }
                finally {
                    gantt.unmaskBody();
                }
            }
            else {
                if (record.meta.modified.id) return;
                let operationSetId = '';
                const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                const projectBucketId = import.meta.env.VITE_MSDYN_PROJECTBUCKET_VALUE;
                const description = 'Create operation set for updating a project task';

                try {
                    gantt.maskBody('Updating task...');
                    operationSetId = await createOperationSet(projectId, description);
                    let msdyn_displaysequence = null;
                    if ('parentIndex' in record.meta.modified && 'orderedParentIndex' in record.meta.modified && Object.keys(record.meta.modified).length === 2) {
                        if (!record.previousSibling) {
                            msdyn_displaysequence = record.nextSibling.msdyn_displaysequence / 2;
                        }
                        if (!record.nextSibling) {
                            msdyn_displaysequence = record.previousSibling.msdyn_displaysequence  + 1;
                        }
                        if (record.previousSibling && record.nextSibling) {
                            msdyn_displaysequence = (record.previousSibling.msdyn_displaysequence + record.nextSibling.msdyn_displaysequence) / 2;
                        }
                    }
                    msdyn_displaysequence += .50001;
                    if (msdyn_displaysequence <= 1) {
                        msdyn_displaysequence = 1.1;
                    }
                    // round to maximum 9 decimal places
                    msdyn_displaysequence = Number(msdyn_displaysequence.toFixed(9));
                    await updateProjectTask(projectId, projectBucketId, operationSetId, record, msdyn_displaysequence);
                    await executeOperationSet(operationSetId);
                    return;
                }
                catch (error) {
                    await abandonOperationSet(operationSetId);
                    console.error('Error:', error);
                }
                finally {
                    gantt.unmaskBody();
                }
            }
        }
        if (action === 'remove') {
            const recordsData = records.map((record) => record.data);
            recordsData.forEach(async(record) => {
                if (record.id.startsWith('_generated')) return;
                let operationSetId = '';
                const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                const description = 'Create operation set for deleting project task';
                try {
                    gantt.maskBody('Deleting task...');
                    operationSetId = await createOperationSet(projectId, description);
                    await deleteProjectTask(operationSetId, record.id);
                    await executeOperationSet(operationSetId);
                    return;
                }
                catch (error) {
                    await abandonOperationSet(operationSetId);
                    console.error('Error:', error);
                }
                finally {
                    gantt.unmaskBody();
                }
            });
        }
    }
    if (storeId === 'dependencies') {
        const recordsData = records.map((record) => record.data);
        recordsData.forEach(async(record) => {
            if (action === 'update') {
                if (`${record.id}`.startsWith('_generated')) {
                    // create new dependency
                    let operationSetId = '';
                    const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                    const description = 'Create operation set for new project task dependency';
                    try {
                        gantt.maskBody('Creating dependency...');
                        operationSetId = await createOperationSet(projectId, description);
                        const createProjectTaskDependencyResponse = await createProjectTaskDependency(projectId, operationSetId, record);
                        await executeOperationSet(operationSetId);
                        // update id
                        gantt.project.dependencyStore.applyChangeset({
                            updated : [
                                // Will set proper id for added task
                                {
                                    $PhantomId : record.id,
                                    id         : JSON.parse(createProjectTaskDependencyResponse.OperationSetResponse)['<OperationSetResponses>k__BackingField'][3].Value
                                }
                            ]
                        });
                        return;
                    }
                    catch (error) {
                        await abandonOperationSet(operationSetId);
                        console.error('Error:', error);
                    }
                    finally {
                        gantt.unmaskBody();
                    }
                }
                else {
                    // 1. delete old dependency
                    let operationSetId = '';
                    let description = '';
                    const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                    description = 'Operation set for updating a project task dependency: delete old and create new';
                    try {
                        gantt.maskBody('Updating dependency...');
                        operationSetId = await createOperationSet(projectId, description);
                        await deleteProjectTaskDependency(record.id, operationSetId);
                        const createProjectTaskDependencyResponse = await createProjectTaskDependency(projectId, operationSetId, record);
                        await executeOperationSet(operationSetId);
                        // update id
                        gantt.project.dependencyStore.applyChangeset({
                            updated : [
                                // Will set proper id for added task
                                {
                                    $PhantomId : record.id,
                                    id         : JSON.parse(createProjectTaskDependencyResponse.OperationSetResponse)['<OperationSetResponses>k__BackingField'][3].Value
                                }
                            ]
                        });
                    }
                    catch (error) {
                        await abandonOperationSet(operationSetId);
                        await abandonOperationSet(operationSetId);
                        console.error('Error:', error);
                    }
                    finally {
                        gantt.unmaskBody();
                    }
                }
            }
            if (action === 'remove') {
                let operationSetId = '';
                const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                const description = 'Create operation set for deleting project task dependency';
                try {
                    gantt.maskBody('Deleting dependency...');
                    operationSetId = await createOperationSet(projectId, description);
                    await deleteProjectTaskDependency(record.id, operationSetId);
                    await executeOperationSet(operationSetId);
                }
                catch (error) {
                    await abandonOperationSet(operationSetId);
                    console.error('Error:', error);
                }
                finally {
                    gantt.unmaskBody();
                }
            }
        });
    }
}
