import * as React from "react";
import GraphAPI from "./GraphAPI";
import { IHelloWorldProps, IHelloWorldWebPartProps } from "./IHelloWorldProps";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export const HelloWorldContext = React.createContext<
  WebPartContext | undefined
>(undefined);
export const HelloWorldProperties = React.createContext<
  IHelloWorldProps | undefined
>(undefined);

const HelloWorld = (props: IHelloWorldWebPartProps) => {
  const { context, properties } = props;

  return (
    <HelloWorldContext.Provider value={context}>
      <HelloWorldProperties.Provider value={properties}>
        <GraphAPI />
      </HelloWorldProperties.Provider>
    </HelloWorldContext.Provider>
  );
};

export default HelloWorld;
