import { ensureScope, getToken } from './auth';

export async function getProjectTasks() {
    try {
        const accessToken = await getToken();
        ensureScope(`https://${
        import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID
      }.api.crm4.dynamics.com/.default`);
        if (!accessToken) {
            throw new Error('Access token is missing');
        }

        const apiUrl = `https://${
        import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID
      }.api.crm4.dynamics.com/api/data/v9.1/msdyn_projecttasks`;

        const response = await fetch(apiUrl, {
            method  : 'GET',
            headers : {
                'Authorization'    : `Bearer ${accessToken}`,
                'OData-MaxVersion' : '4.0',
                'OData-Version'    : '4.0',
                'Accept'           : 'application/json',
                'Content-Type'     : 'application/json'
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        const modData = data.value.filter((item) => item._msdyn_project_value === import.meta.env.VITE_MSDYN_PROJECT_ID);
        return modData;
    }
    catch (error) {
        console.error('Error fetching project tasks:', error);
        throw error;
    }
}

export async function getProjectTaskDependencies() {
    try {
        const accessToken = await getToken();
        ensureScope(`https://${
      import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID
    }.api.crm4.dynamics.com/.default`);
        if (!accessToken) {
            throw new Error('Access token is missing');
        }

        const apiUrl = `https://${
      import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID
    }.api.crm4.dynamics.com/api/data/v9.1/msdyn_projecttaskdependencies`;

        const response = await fetch(apiUrl, {
            method  : 'GET',
            headers : {
                'Authorization'    : `Bearer ${accessToken}`,
                'OData-MaxVersion' : '4.0',
                'OData-Version'    : '4.0',
                'Accept'           : 'application/json',
                'Content-Type'     : 'application/json'
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        const modData = data.value.filter((item) => item._msdyn_project_value === import.meta.env.VITE_MSDYN_PROJECT_ID);
        return modData;
    }
    catch (error) {
        console.error('Error fetching project task dependencies:', error);
        throw error;
    }
}

export async function createProjectTask(projectId, projectBucketId, operationSetId, record, msdyn_displaysequence) {
    const accessToken = await getToken();
    ensureScope(`https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/.default`);
    if (!accessToken) {
        throw new Error('Access token is missing');
    }

    const apiUrl = `https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/api/data/v9.1/msdyn_PssCreateV1`;
    const payload = {
        'Entity' : {
            'msdyn_subject'                  : record.name,
            'msdyn_start'                    : `${record.startDate.toISOString()}`,
            'msdyn_finish'                   : `${record.endDate.toISOString()}`,
            '@odata.type'                    : 'Microsoft.Dynamics.CRM.msdyn_projecttask',
            'msdyn_project@odata.bind'       : `msdyn_projects(${projectId})`,
            'msdyn_projectbucket@odata.bind' : `msdyn_projectbuckets(${projectBucketId})`,
            '_msdyn_parenttask_value'        : record.parentId,
            'msdyn_progress'                 : record.percentDone / 100,
            'msdyn_displaysequence'          : msdyn_displaysequence
        },
        'OperationSetId' : operationSetId
    };

    const response = await fetch(apiUrl, {
        method  : 'POST',
        headers : {
            'Authorization'    : `Bearer ${accessToken}`,
            'OData-MaxVersion' : '4.0',
            'OData-Version'    : '4.0',
            'Accept'           : 'application/json',
            'Content-Type'     : 'application/json'
        },
        body : JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return data;
}

export async function createOperationSet(projectId, description) {
    const accessToken = await getToken();
    ensureScope(`https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/.default`);
    if (!accessToken) {
        throw new Error('Access token is missing');
    }

    const apiUrl = `https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/api/data/v9.1/msdyn_CreateOperationSetV1`;

    const payload = {
        'ProjectId'   : projectId,
        'Description' : description
    };

    const response = await fetch(apiUrl, {
        method  : 'POST',
        headers : {
            'Authorization'    : `Bearer ${accessToken}`,
            'OData-MaxVersion' : '4.0',
            'OData-Version'    : '4.0',
            'Accept'           : 'application/json',
            'Content-Type'     : 'application/json'
        },
        body : JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return data.OperationSetId;
}

export async function executeOperationSet(operationSetId) {
    const accessToken = await getToken();
    ensureScope(`https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/.default`);
    if (!accessToken) {
        throw new Error('Access token is missing');
    }

    const apiUrl = `https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/api/data/v9.1/msdyn_ExecuteOperationSetV1`;

    const payload = {
        'OperationSetId' : operationSetId
    };

    const response = await fetch(apiUrl, {
        method  : 'POST',
        headers : {
            'Authorization'    : `Bearer ${accessToken}`,
            'OData-MaxVersion' : '4.0',
            'OData-Version'    : '4.0',
            'Accept'           : 'application/json',
            'Content-Type'     : 'application/json'
        },
        body : JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return data;
}

export async function abandonOperationSet(operationSetId) {
    const accessToken = await getToken();
    ensureScope(`https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/.default`);
    if (!accessToken) {
        throw new Error('Access token is missing');
    }

    const apiUrl = `https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/api/data/v9.1/msdyn_AbandonOperationSetV1`;

    const payload = {
        'OperationSetId' : operationSetId
    };

    const response = await fetch(apiUrl, {
        method  : 'POST',
        headers : {
            'Authorization'    : `Bearer ${accessToken}`,
            'OData-MaxVersion' : '4.0',
            'OData-Version'    : '4.0',
            'Accept'           : 'application/json',
            'Content-Type'     : 'application/json'
        },
        body : JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return data;
}

export async function updateProjectTask(projectId, projectBucketId, operationSetId, record, msdyn_displaysequence) {
    const accessToken = await getToken();
    ensureScope(`https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/.default`);
    if (!accessToken) {
        throw new Error('Access token is missing');
    }

    const apiUrl = `https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/api/data/v9.1/msdyn_PssUpdateV1`;

    const payloadObj = {
        Entity : {
            msdyn_projecttaskid              : record.id,
            '@odata.type'                    : 'Microsoft.Dynamics.CRM.msdyn_projecttask',
            'msdyn_project@odata.bind'       : `msdyn_projects(${projectId})`,
            'msdyn_projectbucket@odata.bind' : `msdyn_projectbuckets(${projectBucketId})`
        },
        OperationSetId : operationSetId
    };

    if (record.name) {
        payloadObj.Entity.msdyn_subject = record.name;
    }
    if (record.startDate) {
        payloadObj.Entity.msdyn_start = `${record.startDate.toISOString()}`;
    }
    if (record.endDate) {
        payloadObj.Entity.msdyn_finish = `${record.endDate.toISOString()}`;
    }
    if (record.parentId) {
        payloadObj.Entity._msdyn_parenttask_value = record.parentId;
    }
    if (msdyn_displaysequence) {
        payloadObj.Entity.msdyn_displaysequence = msdyn_displaysequence;
    }

    // 	The Progress, EffortCompleted, and EffortRemaining fields can be edited in Project for the Web, but they can't be edited in Project Operations: https://learn.microsoft.com/en-us/dynamics365/project-operations/project-management/schedule-api-preview

    const response = await fetch(apiUrl, {
        method  : 'POST',
        headers : {
            'Authorization'    : `Bearer ${accessToken}`,
            'OData-MaxVersion' : '4.0',
            'OData-Version'    : '4.0',
            'Accept'           : 'application/json',
            'Content-Type'     : 'application/json'
        },
        body : JSON.stringify(payloadObj)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return data;
}

export async function deleteProjectTask(operationSetId, recordId) {
    const accessToken = await getToken();
    ensureScope(`https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/.default`);
    if (!accessToken) {
        throw new Error('Access token is missing');
    }

    const apiUrl = `https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/api/data/v9.1/msdyn_PssDeleteV1`;

    const payload = {
        'EntityLogicalName' : 'msdyn_projecttask',
        'RecordId'          : recordId,
        'OperationSetId'    : operationSetId
    };

    const response = await fetch(apiUrl, {
        method  : 'POST',
        headers : {
            'Authorization'    : `Bearer ${accessToken}`,
            'OData-MaxVersion' : '4.0',
            'OData-Version'    : '4.0',
            'Accept'           : 'application/json',
            'Content-Type'     : 'application/json'
        },
        body : JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return data;
}

export async function createProjectTaskDependency(projectId, operationSetId, record) {
    const accessToken = await getToken();
    ensureScope(`https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/.default`);
    if (!accessToken) {
        throw new Error('Access token is missing');
    }

    const apiUrl = `https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/api/data/v9.1/msdyn_PssCreateV1`;

    const payload = {
        'Entity' : {
            'msdyn_PredecessorTask@odata.bind'    : `msdyn_projecttasks(${record.fromEvent.id ? record.fromEvent.id : record.fromEvent})`,
            'msdyn_SuccessorTask@odata.bind'      : `msdyn_projecttasks(${record.toEvent.id ? record.toEvent.id : record.toEvent})`,
            'msdyn_projecttaskdependencylinktype' : 1,
            '@odata.type'                         : 'Microsoft.Dynamics.CRM.msdyn_projecttaskdependency',
            'msdyn_Project@odata.bind'            : `msdyn_projects(${projectId})`
        },
        'OperationSetId' : operationSetId
    };

    const response = await fetch(apiUrl, {
        method  : 'POST',
        headers : {
            'Authorization'    : `Bearer ${accessToken}`,
            'OData-MaxVersion' : '4.0',
            'OData-Version'    : '4.0',
            'Accept'           : 'application/json',
            'Content-Type'     : 'application/json'
        },
        body : JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return data;
}

export async function deleteProjectTaskDependency(recordId, operationSetId) {
    const accessToken = await getToken();
    ensureScope(`https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/.default`);
    if (!accessToken) {
        throw new Error('Access token is missing');
    }

    const apiUrl = `https://${import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID}.api.crm4.dynamics.com/api/data/v9.1/msdyn_PssDeleteV1`;

    const payload = {
        'EntityLogicalName' : 'msdyn_projecttaskdependency',
        'RecordId'          : recordId,
        'OperationSetId'    : operationSetId
    };

    const response = await fetch(apiUrl, {
        method  : 'POST',
        headers : {
            'Authorization'    : `Bearer ${accessToken}`,
            'OData-MaxVersion' : '4.0',
            'OData-Version'    : '4.0',
            'Accept'           : 'application/json',
            'Content-Type'     : 'application/json'
        },
        body : JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return data;
}
