import { Gantt } from '@bryntum/gantt';
import '@bryntum/gantt/gantt.stockholm.css';
import { signIn } from './auth.js';
import { abandonOperationSet, createOperationSet, createProjectTask, createProjectTaskDependency, deleteProjectTask, deleteProjectTaskDependency, executeOperationSet, getProjectTaskDependencies, getProjectTasks, updateProjectTask, waitForOperationSetCompletion } from './crudFunctions.js';
import CustomTaskModel from './lib/CustomTaskModel.js';

const signInLink = document.getElementById('signin');

let disableCreate = false;
let disableDelete = false;

const gantt = new Gantt({
    appendTo   : 'gantt',
    viewPreset : 'weekAndMonth',
    timeZone   : 'UTC',
    date       : new Date(2024, 10, 1),
    project    : {
        taskModelClass : CustomTaskModel,
        taskStore      : {
            transformFlatData : true,
            modelClass        : CustomTaskModel
        },
        writeAllFields : true
    },
    features : {
        taskMenu : {
            items : {
                // Hide items from the `edit` menu
                copy               : false,
                indent             : false,
                outdent            : false,
                convertToMilestone : false
            }
        },
        taskEdit : {
            items : {
                generalTab : {
                    items : {
                        percentDone : {
                            disabled : true
                        },
                        effort : {
                            disabled : true
                        }
                    }
                },
                resourcesTab : false,
                advancedTab  : false
            }
        }
    },
    listeners : {
        dataChange : function(event) {
            updateMicrosoftProject(event);
        },
        cellClick : function({ target }) {
            if (target.className === 'b-tree-expander b-icon b-icon-tree-collapse' || target.className === 'b-tree-expander b-icon b-icon-tree-expand') {
                disableCreate  = true;
                disableDelete = true;
                setTimeout(() => {
                    disableCreate = false;
                    disableDelete = false;
                }, 50);
            }
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
                effort                : event.msdyn_effort,
                msdyn_displaysequence : event.msdyn_displaysequence,
                manuallyScheduled     : true,
                msdyn_outlinelevel    : event.msdyn_outlinelevel,
                note                  : event.msdyn_descriptionplaintext
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

function calculateNewDisplaySequence(prevSeq, nextSeq) {
    if (prevSeq === undefined) {
        prevSeq = 1;
    }
    if (nextSeq === undefined) {
        return prevSeq + 1;
    }

    let newSeq = (prevSeq + nextSeq) / 2;

    // Round to max 9 decimal places
    // so we don’t exceed the msdyn_displaysequence column’s precision limit
    const seqParts = newSeq.toString().split('.');
    const decimalPart = seqParts[1] || '';
    const decimalCount = decimalPart.length;
    if (decimalCount > 9) {
        // Round to 9 decimals
        newSeq = Math.round(newSeq * 1e9) / 1e9;
    }
    return newSeq;
}

function getSubtaskBoundaries(parentTask, record) {
    let prevSeq, nextSeq;
    // 1. Figure out prev boundary
    if (record.previousSibling) {
        prevSeq = record.previousSibling.msdyn_displaysequence;
    }
    else {
        // No previous sibling => use parent's sequence
        prevSeq = parentTask.data.msdyn_displaysequence;
    }

    // 2. Figure out next boundary
    if (record.nextSibling) {
        nextSeq = record.nextSibling.msdyn_displaysequence;
    }
    else {
        // If parent has a next sibling at top level, use that
        if (parentTask.nextSibling) {
            nextSeq = parentTask.nextSibling.data.msdyn_displaysequence;
        }
        else {
            // Parent is last => fallback
            nextSeq = prevSeq + 1;
        }
    }

    return { prevSeq, nextSeq };
}

async function updateMicrosoftProject({ action, record, store, records }) {
    const storeId = store.id;
    if (storeId === 'tasks') {
        if (action === 'update') {
            if (`${record.id}`.startsWith('_generated')) {
                if (disableCreate) return;
                if (!record.name) return;
                let operationSetId = '';
                const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                const projectBucketId = import.meta.env.VITE_MSDYN_PROJECTBUCKET_VALUE;
                const description = 'Create operation set for new project task';

                try {
                    gantt.maskBody('Creating task...');
                    operationSetId = await createOperationSet(projectId, description);

                    let previousSibling = record.previousSibling?.msdyn_displaysequence;
                    let nextSibling = record.nextSibling?.msdyn_displaysequence;

                    // check if subtask
                    const isSubtask = record.parentId !== null;

                    if (isSubtask) {
                        const parentTask = gantt.taskStore.getById(record.parentId);
                        const { prevSeq, nextSeq } = getSubtaskBoundaries(parentTask, record);
                        previousSibling = prevSeq;
                        nextSibling = nextSeq;
                    }
                    // if previous sibling has children, get the last child's display sequence
                    if (previousSibling && record.previousSibling?.children?.length > 0) {
                        previousSibling = record.previousSibling.children.map((child) => child.data.msdyn_displaysequence).sort((a, b) => a - b).at(-1);
                    }
                    // prev and no next -  check if previous sibling has children - if yes -> get its next sibling's display sequence
                    if (previousSibling && record.previousSibling?.children?.length > 0 && !nextSibling) {
                        const newNextSibling = record.previousSibling.nextSibling;
                        if (newNextSibling) {
                            nextSibling = newNextSibling.data.msdyn_displaysequence;
                        }
                    }
                    const msdyn_displaysequence = calculateNewDisplaySequence(previousSibling, nextSibling);
                    const createProjectTaskResponse = await createProjectTask(projectId, projectBucketId, operationSetId, record, msdyn_displaysequence);
                    const newId = JSON.parse(createProjectTaskResponse.OperationSetResponse)['<OperationSetResponses>k__BackingField'][3].Value;

                    await executeOperationSet(operationSetId);
                    // update id
                    gantt.project.taskStore.applyChangeset({
                        updated : [
                        // Will set proper id for added task
                            {
                                $PhantomId : record.id,
                                id         : newId
                            }
                        ]
                    });
                    // check if task available for CRUD operations
                    await waitForOperationSetCompletion(operationSetId, 'task');
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
                if (Object.keys(record.meta.modified).length === 0) return;
                if (record.meta.modified.effort === 0 && Object.keys(record.meta.modified).length === 1) return;
                if (record.meta.modified.id && Object.keys(record.meta.modified).length === 1) return;
                let operationSetId = '';
                const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                const description = 'Create operation set for updating a project task';

                try {
                    operationSetId = await createOperationSet(projectId, description);

                    let previousSibling = record.previousSibling?.msdyn_displaysequence;
                    let nextSibling = record.nextSibling?.msdyn_displaysequence;

                    // check if subtask
                    const isSubtask = record.parentId !== null;

                    if (isSubtask) {
                        const parentTask = gantt.taskStore.getById(record.parentId);
                        const { prevSeq, nextSeq } = getSubtaskBoundaries(parentTask, record);
                        previousSibling = prevSeq;
                        nextSibling = nextSeq;
                    }
                    // if previous sibling has children, get the last child's display sequence
                    if (previousSibling && record.previousSibling?.children?.length > 0) {
                        previousSibling = record.previousSibling.children.map((child) => child.data.msdyn_displaysequence).sort((a, b) => a - b).at(-1);
                    }
                    // prev and no next -  check if previous sibling has children - if yes -> get its next sibling's display sequence
                    if (previousSibling && record.previousSibling?.children?.length > 0 && !nextSibling) {
                        const newNextSibling = record.previousSibling.nextSibling;
                        if (newNextSibling) {
                            nextSibling = newNextSibling.data.msdyn_displaysequence;
                        }
                    }
                    const msdyn_displaysequence = calculateNewDisplaySequence(previousSibling, nextSibling);
                    const isReorder = record.meta.modified.orderedParentIndex !== undefined;
                    gantt.project.taskStore.commit();
                    const isParentTask = record.children?.length > 0;
                    await updateProjectTask(operationSetId, record, msdyn_displaysequence, isReorder, isParentTask);
                    await executeOperationSet(operationSetId);
                    return;
                }
                catch (error) {
                    await abandonOperationSet(operationSetId);
                    console.error('Error:', error);
                }
            }
        }
        if (action === 'remove') {
            if (disableDelete) return;
            const recordsData = records.map((record) => record.data);
            recordsData.forEach(async(record) => {
                if (record.id.startsWith('_generated')) return;
                let operationSetId = '';
                const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                const description = 'Create operation set for deleting project task';
                try {
                    operationSetId = await createOperationSet(projectId, description);
                    await deleteProjectTask(operationSetId, record.id);
                    await executeOperationSet(operationSetId);
                    return;
                }
                catch (error) {
                    await abandonOperationSet(operationSetId);
                    console.error('Error:', error);
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
                        await waitForOperationSetCompletion(operationSetId, 'dependency');
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
                        await waitForOperationSetCompletion(operationSetId, 'dependency');
                        return;
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
                    operationSetId = await createOperationSet(projectId, description);
                    await deleteProjectTaskDependency(record.id, operationSetId);
                    await executeOperationSet(operationSetId);
                }
                catch (error) {
                    await abandonOperationSet(operationSetId);
                    console.error('Error:', error);
                }
            }
        });
    }
}
