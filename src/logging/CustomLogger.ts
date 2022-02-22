import ICustomLogger from './ICustomLogger';
import ICustomLogMessage from './ICustomLogMessage';
import { sp } from '@pnp/sp';
 
export class CustomLogger implements ICustomLogger {
 
    public Log = async (logMessage: ICustomLogMessage) => {
        try {
            this.saveLogs(logMessage, "Log");
        } catch (error) {
            //Can't do anything
            console.error(error.Message);
        }
    }
 
    public Warn = async (logMessage: ICustomLogMessage) => {
        try {
            this.saveLogs(logMessage, "Warn");
        } catch (error) {
            //Can't do anything
            console.error(error.Message);
        }
    }
 
    public Verbose = async (logMessage: ICustomLogMessage) => {
        try {
            this.saveLogs(logMessage, "Verbose");
        } catch (error) {
            //Can't do anything
            console.error(error.Message);
        }
    }
 
    public Error = async (logMessage: ICustomLogMessage) => {
        try {
            console.error(logMessage.Message);
            this.saveLogs(logMessage, "Error");
        } catch (error) {
            //Can't do anything
            console.error(error.Message);
        }
    }
 
    private saveLogs = async (logMessage: ICustomLogMessage, logType: string) => {
        var today = new Date();
        sp.web.lists.getByTitle('DailyUpdateWebpartLogs').items.add({
            Title: logMessage.ComponentName,
            MethodName: logMessage.MethodName,
            Message: logMessage.Message,
            LogType: logType,
            Date: today.toDateString() + " " + today.toLocaleTimeString()
        });
    }
}
 
const customLogger = new CustomLogger();
export default customLogger;
