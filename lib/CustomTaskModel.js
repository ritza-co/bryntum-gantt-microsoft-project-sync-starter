import { TaskModel } from '@bryntum/gantt';

// Custom event model
export default class CustomTaskModel extends TaskModel {
    static $name = 'CustomTaskModel';
    static fields = [
        { name : 'description', type : 'string' },
        { name : 'outlineLevel', type : 'number' }
    ];
}
