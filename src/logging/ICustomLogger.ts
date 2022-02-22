import ICustomLogMessage from './ICustomLogMessage';
 
interface ICustomLogger {
    Log(logMessage: ICustomLogMessage): void;
    Warn(logMessage: ICustomLogMessage): void;
    Verbose(logMessage: ICustomLogMessage): void;
    Error(logMessage: ICustomLogMessage): void;
}
 
export default ICustomLogger;
