import { Subject } from "rx-lite";

export class RxJsEventEmitter {
  public subjects: Object;
  public readonly hasOwnProp: any = {}.hasOwnProperty;

  private constructor() {

    this.subjects = {};
  }

  // tslint:disable:no-string-literal
  /**
   * Singleton for the page so we capture all the Observers and Observables in one global array;
   */
  public static getInstance(): RxJsEventEmitter {

    if (!(window as any)["RxJsEventEmitter"]) {

      (window as any)["RxJsEventEmitter"] = new RxJsEventEmitter();
    }
    return (window as any)["RxJsEventEmitter"];
  }

  /**
   * Emitts (broadcasts) event to Observers (Subscribers).
   * @param name name of the event
   * @param data event data
   */
  public emit(name: string, data: Object): void {
    let fnName: string = this._createName(name);

    if (!(this.subjects as any)[fnName]) {

      (this.subjects as any)[fnName] = new Subject();
    }

    (this.subjects as any)[fnName].onNext(data);
  }

  /**
   * Subscribes for event stream.
   * If the event is broadcasted then handler (method)
   * would be triggered and would receive data from the broadcasted event as method param.
   * @param name name of the event
   * @param handler event handler (method)
   */
  public on(name: string, handler: any): void {
    let fnName: string = this._createName(name);

    if (!(this.subjects as any)[fnName]) {

      (this.subjects as any)[fnName] = new Subject();
    }

    (this.subjects as any)[fnName].subscribe(handler);
  }

  /**
   * Unsubscribes Observer (Subscriber) from event.
   * @param name name of the event
   */
  public off(name: string): void {
    let fnName: string = this._createName(name);

    if ((this.subjects as any)[fnName]) {

      (this.subjects as any)[fnName].dispose();
      delete (this.subjects as any)[fnName];

    }

  }

  /**
   * Not tested.
   */
  public dispose(): void {

    let subjects: Object = this.subjects;

    for (let prop in subjects) {
      if (this.hasOwnProp.call(subjects, prop)) {
        (subjects as any)[prop].dispose();
      }
    }

    this.subjects = {};
  }

  private _createName(name: string): string {
    return `$${name}`;
  }

}