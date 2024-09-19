import { Gantt } from '@bryntum/gantt';
import '@bryntum/gantt/gantt.stockholm.css';
import { signIn } from './auth.js';
import { abandonOperationSet, createOperationSet, createProjectTask, createProjectTaskDependency, deleteProjectTask, deleteProjectTaskDependency, executeOperationSet, getProjectTaskDependencies, getProjectTasks, updateProjectTask } from './crudFunctions.js';
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
                id                : event.msdyn_projecttaskid,
                parentId          : event._msdyn_parenttask_value,
                name              : event.msdyn_subject,
                startDate         : startDateLocal,
                endDate           : finishDateLocal,
                percentDone       : event.msdyn_progress * 100,
                parentIndex       : parseInt(event.msdyn_displaysequence) + 1,
                manuallyScheduled : true,
                outlineLevel      : event.msdyn_outlinelevel
            });
        });
        ganttTasks.sort((a, b) => a.parentIndex - b.parentIndex);
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
                    operationSetId = await createOperationSet(projectId, description);
                    const createProjectTaskResponse = await createProjectTask(projectId, projectBucketId, operationSetId, record);
                    await executeOperationSet(operationSetId);
                    // update id
                    gantt.project.taskStore.applyChangeset({
                        updated : [
                            // Will set proper id for added task
                            {
                                $PhantomId : record.id,
                                id         : createProjectTaskResponse['<OperationSetResponses>k__BackingField'][3].Value
                            }
                        ]
                    });
                    return;
                }
                catch (error) {
                    await abandonOperationSet(operationSetId);
                    console.error('Error:', error);
                }
            }
            else {
                let operationSetId = '';
                const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                const projectBucketId = import.meta.env.VITE_MSDYN_PROJECTBUCKET_VALUE;
                const description = 'Create operation set for updating a project task';

                try {
                    operationSetId = await createOperationSet(projectId, description);
                    await updateProjectTask(projectId, projectBucketId, operationSetId, record);
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
                }
                else {
                    // 1. delete old dependency
                    let operationSetId = '';
                    let description = '';
                    const projectId = import.meta.env.VITE_MSDYN_PROJECT_ID;
                    description = 'Operation set for updating a project task dependency: delete old and create new';
                    try {
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
