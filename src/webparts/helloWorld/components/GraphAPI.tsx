import * as React from "react";
import { HelloWorldContext, HelloWorldProperties } from "./HelloWorld";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import "./GraphAPI.css";

import { TextField, Button } from "@mui/material";

interface ISiteCollection {
  id: string;
  webUrl: string;
  displayName: string;
  siteCollection: {
    hostname: string;
  };
}

export const InputListName = ({
  inputData,
  setInputData,
}: {
  inputData: string;
  setInputData: React.Dispatch<React.SetStateAction<string>>;
}) => {
  const handleSubmit = (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
  };

  const handleChange = (
    e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>
  ) => {
    setInputData(e.target.value);
  };

  return (
    <form onSubmit={(e) => handleSubmit(e)}>
      <TextField
        id="outlined-basic"
        label="Enter list name"
        variant="outlined"
        margin="normal"
        value={inputData}
        onChange={(e) => handleChange(e)}
      />
    </form>
  );
};

export const CreateList = ({
  inputData,
  siteID,
}: {
  inputData: string;
  siteID: string | undefined;
}) => {
  const context = React.useContext(HelloWorldContext);

  const [isListCreationSucceed, setIsListCreationSucceed] =
    React.useState<boolean>(false);

  const createList = async () => {
    try {
      const body = {
        displayName: inputData,
        columns: [
          // {
          //   name: "Author",
          //   text: {},
          // },
        ],
        list: {
          template: "genericList",
        },
      };

      const graphClient: MSGraphClientV3 | undefined =
        await context?.msGraphClientFactory.getClient("3");

      await graphClient?.api(`/sites/${siteID}/lists`).post(body);

      setIsListCreationSucceed(true);
    } catch (err) {
      console.log("Error: ", err);
    }
  };

  return (
    <>
      <div className="button">
        <Button variant="contained" size="medium" onClick={createList}>
          Create List
        </Button>
        {isListCreationSucceed ? (<div>{inputData} was succesfully created!!</div>) : ""}
      </div>
    </>
  );
};

const GraphAPI = () => {
  const context = React.useContext(HelloWorldContext);
  const properties = React.useContext(HelloWorldProperties);

  const [multipleSiteCollectionData, setMultipleSiteCollectionData] =
    React.useState<ISiteCollection[]>([]);

  const [siteCollectionData, setSiteCollectionData] =
    React.useState<ISiteCollection>();

  const [myName, setMyName] = React.useState<string | null | undefined>("");

  const [inputData, setInputData] = React.useState<string>("");

  React.useEffect(() => {
    try {
      const fetchData = async () => {
        const graphClient: MSGraphClientV3 | undefined =
          await context?.msGraphClientFactory.getClient("3");
        const siteCollectionResponse = await graphClient
          ?.api("/sites?search=*")
          .get();
        const siteCollectionData = siteCollectionResponse.value;
        setMultipleSiteCollectionData(siteCollectionData);
      };
      fetchData();
    } catch (err) {
      console.log("Error: ", err);
    }
  }, []);

  const getUserInfo = () => {
    context?.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api("/me")
          .get(
            (error: any, response: MicrosoftGraph.User, rawResponse?: any) => {
              setMyName(response.displayName);
            }
          );
      })
      .catch((err) => {
        console.log("erris", err);
      });
  };

  const getSiteInfo = async () => {
    try {
      const graphClient: MSGraphClientV3 | undefined =
        await context?.msGraphClientFactory.getClient("3");

      const siteCollectionResponse = await graphClient
        ?.api(`/sites?search=${context?.pageContext.web.title}`)
        .get();

      setSiteCollectionData(siteCollectionResponse.value[0]);
    } catch (err) {
      console.log("Error: ", err);
    }
  };

  return (
    <>
      <InputListName inputData={inputData} setInputData={setInputData} />
      <CreateList inputData={inputData} siteID={siteCollectionData?.id} />
      <hr />
      <h3>Properties</h3>
      <div>description: {properties?.description}</div>
      <div>test: {properties?.test}</div>
      <hr />
      <div>
        <button type="button" onClick={getUserInfo}>
          Get My Info
        </button>
        <h3>My Name is {myName}</h3>
      </div>
      <hr />
      <div>
        <button type="button" onClick={getSiteInfo}>
          Get This Site's Info
        </button>
        <h3>This site info:</h3>
        <ul className="ThisSiteInfo">
          <li>DisplayName: {siteCollectionData?.displayName}</li>
          <li>Site ID: {siteCollectionData?.id}</li>
          <li>Site URL: {siteCollectionData?.webUrl}</li>
        </ul>
      </div>
      <hr />
      <h3>A site info:</h3>
      <div>
        <ul className="SiteInfo">
          {multipleSiteCollectionData
            .filter((value, index) => {
              return index === 3;
            })
            .map((site: ISiteCollection) => {
              return (
                <li key={site.id}>
                  <div>DisplayName: {site.displayName}</div>
                  <div>Site ID: {site.id}</div>
                  <div>Site URL: {site.webUrl}</div>
                </li>
              );
            })}
        </ul>
      </div>
    </>
  );
};

export default GraphAPI;
