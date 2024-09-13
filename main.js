import { Gantt } from '@bryntum/gantt';
import '@bryntum/gantt/gantt.stockholm.css';

const gantt = new Gantt({
    appendTo   : 'gantt',
    viewPreset : 'weekAndMonth',
    timeZone   : 'UTC',
    date       : new Date(2024, 10, 1),
    project    : {
        tasksData : [

            { id : 1, name : 'Create docs', startDate : '2024-10-07', endDate : '2024-10-18' },
            { id : 2, name : 'Write guides', startDate : '2024-10-21', endDate : '2024-10-31' }
        ],
        dependenciesData : [
            { fromTask : 1, toTask : 2 }
        ]
    }
});
