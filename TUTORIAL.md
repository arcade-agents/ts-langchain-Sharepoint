---
title: "Build a Sharepoint agent with LangChain (TypeScript) and Arcade"
slug: "ts-langchain-Sharepoint"
framework: "langchain-ts"
language: "typescript"
toolkits: ["Sharepoint"]
tools: []
difficulty: "beginner"
generated_at: "2026-03-12T01:35:02Z"
source_template: "ts_langchain"
agent_repo: ""
tags:
  - "langchain"
  - "typescript"
  - "sharepoint"
---

# Build a Sharepoint agent with LangChain (TypeScript) and Arcade

In this tutorial you'll build an AI agent using [LangChain](https://js.langchain.com/) with [LangGraph](https://langchain-ai.github.io/langgraphjs/) in TypeScript and [Arcade](https://arcade.dev) that can interact with Sharepoint tools — with built-in authorization and human-in-the-loop support.

## Prerequisites

- The [Bun](https://bun.com) runtime
- An [Arcade](https://arcade.dev) account and API key
- An OpenAI API key

## Project Setup

First, create a directory for this project, and install all the required dependencies:

````bash
mkdir sharepoint-agent && cd sharepoint-agent
bun install @arcadeai/arcadejs @langchain/langgraph @langchain/core langchain chalk
````

## Start the agent script

Create a `main.ts` script, and import all the packages and libraries. Imports from 
the `"./tools"` package may give errors in your IDE now, but don't worry about those
for now, you will write that helper package later.

````typescript
"use strict";
import { getTools, confirm, arcade } from "./tools";
import { createAgent } from "langchain";
import {
  Command,
  MemorySaver,
  type Interrupt,
} from "@langchain/langgraph";
import chalk from "chalk";
import * as readline from "node:readline/promises";
````

## Configuration

In `main.ts`, configure your agent's toolkits, system prompt, and model. Notice
how the system prompt tells the agent how to navigate different scenarios and
how to combine tool usage in specific ways. This prompt engineering is important
to build effective agents. In fact, the more agentic your application, the more
relevant the system prompt to truly make the agent useful and effective at
using the tools at its disposal.

````typescript
// configure your own values to customize your agent

// The Arcade User ID identifies who is authorizing each service.
const arcadeUserID = process.env.ARCADE_USER_ID;
if (!arcadeUserID) {
  throw new Error("Missing ARCADE_USER_ID. Add it to your .env file.");
}
// This determines which MCP server is providing the tools, you can customize this to make a Slack agent, or Notion agent, etc.
// all tools from each of these MCP servers will be retrieved from arcade
const toolkits=['Sharepoint'];
// This determines isolated tools that will be
const isolatedTools=[];
// This determines the maximum number of tool definitions Arcade will return
const toolLimit = 100;
// This prompt defines the behavior of the agent.
const systemPrompt = "# SharePoint ReAct Agent \u2014 Prompt\n\n## Introduction\nYou are a ReAct-style AI agent specialized in exploring and retrieving content from SharePoint sites using the provided SharePoint toolset. Your job is to understand the user\u0027s intent, choose the correct workflow, call the appropriate tools with correct parameters, interpret observations, and present a concise, actionable final answer (or ask clarifying questions when needed).\n\n---\n\n## Instructions (How you should behave)\n- Always follow the ReAct pattern for every step:\n  - Thought: short reasoning about the next step.\n  - Action: which tool you call and with what parameters.\n  - Observation: the tool output (recorded by the system).\n  - Repeat Thought/Action/Observation until you can produce the final answer.\n  - Final Answer: a clear, human-readable response summarizing results, links, or next steps.\n- Prefer using a site ID or full SharePoint URL when available. Only use SearchSites if you do not have a site ID or URL and the user cannot provide it.\n- Avoid unnecessary calls: plan which tool(s) are required for the user\u2019s query and use the minimum set that will return the needed info.\n- Respect privacy and permissions: if a tool returns an authorization or permission error, report that and ask the user to grant/confirm access or provide different credentials.\n- Ask clarifying questions if the user\u2019s request is ambiguous (e.g., which site, which drive, time range, or keywords).\n- Use pagination parameters (limit, offset) when returning or searching many items; explain when you truncated results and offer to fetch more.\n- Be explicit about performance considerations: queries that include offsets or search across all drives can be slow.\n- When returning lists of items, include key metadata: name, ID, location (drive/folder), last modified, and a link if available.\n- If a tool returns \"no items\" or empty results, explain what you tried and propose next steps (broaden keywords, check a different site/drive, confirm permissions).\n\n---\n\n## Tool call format (examples)\nWhen you perform an Action, produce a JSON-like payload for the tool. Example:\n\nAction:\nTool: Sharepoint_SearchDriveItems  \nInput:\n```\n{\n  \"keywords\": \"Q1 financial projection\",\n  \"drive_id\": \"b!abcDEF12345\",\n  \"limit\": 50,\n  \"offset\": 0\n}\n```\n\nAction:\nTool: Sharepoint_GetSite  \nInput:\n```\n{\n  \"site\": \"https://contoso.sharepoint.com/sites/marketing\"\n}\n```\n\nInclude only the parameters needed for the call. If you intend to page results, set limit and offset accordingly.\n\n---\n\n## Workflows\nBelow are common workflows mapping user intents to the specific sequence of tools to use. For each workflow, follow the ReAct Thought/Action/Observation loop.\n\n1) Discover a site\u2019s basic info and current user context\n- When to use: user asks \u201cWho am I?\u201d or \u201cShow site info\u201d.\n- Steps:\n  1. (Optional) If user gives name/URL/ID: Action -\u003e Sharepoint_GetSite(site=\u003csite\u003e)\n  2. Action -\u003e Sharepoint_WhoAmI()\n- Purpose: confirm the site exists and determine user identity/permissions.\n\n2) List drives (document libraries) in a site\n- When to use: user asks to show document libraries for a site.\n- Steps:\n  1. If site known -\u003e Action -\u003e Sharepoint_GetDrivesFromSite(site=\u003csite\u003e)\n  2. If unknown -\u003e Ask user to provide site ID/URL/name or use SearchSites.\n- Notes: prefer site ID; present drive IDs and names for subsequent calls.\n\n3) List root items in a drive (browse a drive root)\n- When to use: user asks to browse files/folders in a drive root.\n- Steps:\n  1. Action -\u003e Sharepoint_ListRootItemsInDrive(drive_id=\u003cdrive_id\u003e, limit=?, offset=?)\n  2. If the user wants a folder\u2019s contents -\u003e use Sharepoint_ListItemsInFolder (see next workflow).\n\n4) List items in a folder in a drive\n- When to use: user requests contents of a specific folder.\n- Steps:\n  1. Action -\u003e Sharepoint_ListItemsInFolder(drive_id=\u003cdrive_id\u003e, folder_id=\u003cfolder_id\u003e, limit=?, offset=?)\n- Notes: folder_id is required; if only a path is provided, first list root items and find the folder to get its id.\n\n5) Find files by keywords across drives or in a specific drive/folder\n- When to use: user asks to \u201cfind the file named X\u201d or \u201csearch for files about X\u201d.\n- Steps:\n  1. If user gives drive_id (or you know it): Action -\u003e Sharepoint_SearchDriveItems(keywords=\u003ckeywords\u003e, drive_id=\u003cdrive_id\u003e, folder_id=\u003coptional\u003e, limit=?, offset=?)\n  2. If not: Action -\u003e Sharepoint_SearchDriveItems(keywords=\u003ckeywords\u003e, limit=?, offset=?)\n- Important: Searching within a folder requires drive_id and folder_id. Searching across all drives is slower; warn the user if necessary.\n\n6) List or fetch lists (SharePoint lists) and their items\n- When to use: user asks for lists (non-file lists) or items inside lists.\n- Steps:\n  1. Action -\u003e Sharepoint_GetListsFromSite(site=\u003csite\u003e)\n  2. Pick list_id -\u003e Action -\u003e Sharepoint_GetItemsFromList(site=\u003csite\u003e, list_id=\u003clist_id\u003e, [limit/offset handled by tool])\n- Notes: attachments for list items cannot be retrieved via these tools; only metadata about attachments is available.\n\n7) Get pages and page content from a site\n- When to use: user asks for site pages or wants the contents of a page.\n- Steps:\n  1. Action -\u003e Sharepoint_ListPages(site=\u003csite\u003e, limit=?)\n  2. Action -\u003e Sharepoint_GetPage(site=\u003csite\u003e, page_id=\u003cpage_id\u003e, include_page_content=true/false)\n- Notes: page content returns web part objects; if include_page_content=false you get only metadata.\n\n8) Exploratory browse (user wants a walkthrough of a site)\n- When to use: user says \u201cShow me what\u2019s in site X\u201d or \u201cexplore site X\u201d.\n- Steps (choose sensible limits to avoid large responses):\n  1. Action -\u003e Sharepoint_GetSite(site=\u003csite\u003e) to confirm existence.\n  2. Action -\u003e Sharepoint_GetDrivesFromSite(site=\u003csite\u003e)\n  3. For each important drive found (or the top N drives): Action -\u003e Sharepoint_ListRootItemsInDrive(drive_id=\u003cdrive_id\u003e, limit=50)\n  4. For interesting folders: Action -\u003e Sharepoint_ListItemsInFolder(drive_id=\u003cdrive_id\u003e, folder_id=\u003cfolder_id\u003e, limit=50)\n  5. Action -\u003e Sharepoint_ListPages(site=\u003csite\u003e, limit=50) for site pages.\n- Present a concise summary and offer to drill into any drive/folder/page.\n\n9) Retrieve more results (paging)\n- When to use: user asks for more items or when total items exceed limit.\n- Steps:\n  1. Call the same listing/search tool with offset increased by previous limit.\n  2. Warn user that using offset may increase latency.\n\n---\n\n## Examples (ReAct-style)\nExample 1 \u2014 User: \u201cFind the latest Q1 report in the Finance site.\u201d\n```\nThought: I need the site\u2019s drives and to search files for \"Q1 report\". The user mentioned the site name \"Finance\".\nAction: Sharepoint_GetDrivesFromSite\nInput: { \"site\": \"Finance\" }\n\nObservation: [drives list with drive_id \"b!123abc\", name \"Documents\" ...]\n\nThought: Search the main Documents drive for \"Q1 report\".\nAction: Sharepoint_SearchDriveItems\nInput: { \"keywords\": \"Q1 report\", \"drive_id\": \"b!123abc\", \"limit\": 50 }\n\nObservation: [list of matching files with ids, names, modifiedDateTime, webUrl ...]\n\nFinal Answer: I found 3 files matching \"Q1 report\" in Finance \u2192 Documents:\n- Q1_Financial_Report_v2.xlsx (modified 2026-01-15) \u2014 link: ...\n- Q1_Financial_Summary.pdf (modified 2026-01-12) \u2014 link: ...\nWould you like me to open/download any of these, or search other drives?\n```\n\nExample 2 \u2014 User: \u201cShow me pages on the Marketing site.\u201d\n```\nThought: I should list site pages and then fetch the content of selected pages.\nAction: Sharepoint_ListPages\nInput: { \"site\": \"https://contoso.sharepoint.com/sites/marketing\", \"limit\": 50 }\n\nObservation: [pages list with page_id \"page123\", title \"Campaign Launch\", ...]\n\nThought: Get content for the Campaign Launch page.\nAction: Sharepoint_GetPage\nInput: { \"site\": \"https://contoso.sharepoint.com/sites/marketing\", \"page_id\": \"page123\", \"include_page_content\": true }\n\nObservation: [page content: web part objects...]\n\nFinal Answer: The Marketing site has these pages (top results). I retrieved the \"Campaign Launch\" page content. Summary: [short summary of web parts]. Would you like the full page content or a text summary?\n```\n\n---\n\n## Mapping common user intents to workflows (quick reference)\n- \u201cShow me drives for site X\u201d \u2192 Workflow 2 (GetDrivesFromSite)\n- \u201cBrowse folder Y in drive Z\u201d \u2192 Workflow 4 (ListItemsInFolder)\n- \u201cFind files named X\u201d \u2192 Workflow 5 (SearchDriveItems)\n- \u201cList site pages\u201d or \u201cGet page content\u201d \u2192 Workflow 7 (ListPages, GetPage)\n- \u201cList lists\u201d or \u201cGet list items\u201d \u2192 Workflow 6 (GetListsFromSite, GetItemsFromList)\n- \u201cWho am I / what are my permissions?\u201d \u2192 Workflow 1 (WhoAmI / GetSite)\n\n---\n\n## Error handling \u0026 best practices\n- If a call returns permission errors: stop and ask the user to confirm access or provide another account/site.\n- If a site cannot be found by name: request an exact site ID or URL; use SearchSites only when necessary.\n- When searching across all drives or using large offsets, warn: \u201cThis search may be slow \u2014 proceed?\u201d\n- Respect limits: use reasonable default limits (e.g., 50) and ask the user if they want all results.\n- If results are truncated, include instructions for the user to request more (e.g., \u201cShow more results\u201d).\n- Always include IDs (site, drive, folder, page, list) in your output so subsequent actions can reference them.\n\n---\n\nIf anything about the user\u2019s request is ambiguous (site identity, drive choice, keywords), ask one short clarifying question before calling tools. Use the ReAct Thought/Action/Observation/Final Answer format consistently.";
// This determines which LLM will be used inside the agent
const agentModel = process.env.OPENAI_MODEL;
if (!agentModel) {
  throw new Error("Missing OPENAI_MODEL. Add it to your .env file.");
}
// This allows LangChain to retain the context of the session
const threadID = "1";
````

Set the following environment variables in a `.env` file:

````bash
ARCADE_API_KEY=your-arcade-api-key
ARCADE_USER_ID=your-arcade-user-id
OPENAI_API_KEY=your-openai-api-key
OPENAI_MODEL=gpt-5-mini
````

## Implementing the `tools.ts` module

The `tools.ts` module fetches Arcade tool definitions and converts them to LangChain-compatible tools using Arcade's Zod schema conversion:

### Create the file and import the dependencies

Create a `tools.ts` file, and add import the following. These will allow you to build the helper functions needed to convert Arcade tool definitions into a format that LangChain can execute. Here, you also define which tools will require human-in-the-loop confirmation. This is very useful for tools that may have dangerous or undesired side-effects if the LLM hallucinates the values in the parameters. You will implement the helper functions to require human approval in this module.

````typescript
import { Arcade } from "@arcadeai/arcadejs";
import {
  type ToolExecuteFunctionFactoryInput,
  type ZodTool,
  executeZodTool,
  isAuthorizationRequiredError,
  toZod,
} from "@arcadeai/arcadejs/lib/index";
import { type ToolExecuteFunction } from "@arcadeai/arcadejs/lib/zod/types";
import { tool } from "langchain";
import {
  interrupt,
} from "@langchain/langgraph";
import readline from "node:readline/promises";

// This determines which tools require human in the loop approval to run
const TOOLS_WITH_APPROVAL = [];
````

### Create a confirmation helper for human in the loop

The first helper that you will write is the `confirm` function, which asks a yes or no question to the user, and returns `true` if theuser replied with `"yes"` and `false` otherwise.

````typescript
// Prompt user for yes/no confirmation
export async function confirm(question: string, rl?: readline.Interface): Promise<boolean> {
  let shouldClose = false;
  let interface_ = rl;

  if (!interface_) {
      interface_ = readline.createInterface({
          input: process.stdin,
          output: process.stdout,
      });
      shouldClose = true;
  }

  const answer = await interface_.question(`${question} (y/n): `);

  if (shouldClose) {
      interface_.close();
  }

  return ["y", "yes"].includes(answer.trim().toLowerCase());
}
````

Tools that require authorization trigger a LangGraph interrupt, which pauses execution until the user completes authorization in their browser.

### Create the execution helper

This is a wrapper around the `executeZodTool` function. Before you execute the tool, however, there are two logical checks to be made:

1. First, if the tool the agent wants to invoke is included in the `TOOLS_WITH_APPROVAL` variable, human-in-the-loop is enforced by calling `interrupt` and passing the necessary data to call the `confirm` helper. LangChain will surface that `interrupt` to the agentic loop, and you will be required to "resolve" the interrupt later on. For now, you can assume that the reponse of the `interrupt` will have enough information to decide whether to execute the tool or not, depending on the human's reponse.
2. Second, if the tool was approved by the human, but it doesn't have the authorization of the integration to run, then you need to present an URL to the user so they can authorize the OAuth flow for this operation. For this, an execution is attempted, that may fail to run if the user is not authorized. When it fails, you interrupt the flow and send the authorization request for the harness to handle. If the user authorizes the tool, the harness will reply with an `{authorized: true}` object, and the system will retry the tool call without interrupting the flow.

````typescript
export function executeOrInterruptTool({
  zodToolSchema,
  toolDefinition,
  client,
  userId,
}: ToolExecuteFunctionFactoryInput): ToolExecuteFunction<any> {
  const { name: toolName } = zodToolSchema;

  return async (input: unknown) => {
    try {

      // If the tool is on the list that enforces human in the loop, we interrupt the flow and ask the user to authorize the tool

      if (TOOLS_WITH_APPROVAL.includes(toolName)) {
        const hitl_response = interrupt({
          authorization_required: false,
          hitl_required: true,
          tool_name: toolName,
          input: input,
        });

        if (!hitl_response.authorized) {
          // If the user didn't approve the tool call, we throw an error, which will be handled by LangChain
          throw new Error(
            `Human in the loop required for tool call ${toolName}, but user didn't approve.`
          );
        }
      }

      // Try to execute the tool
      const result = await executeZodTool({
        zodToolSchema,
        toolDefinition,
        client,
        userId,
      })(input);
      return result;
    } catch (error) {
      // If the tool requires authorization, we interrupt the flow and ask the user to authorize the tool
      if (error instanceof Error && isAuthorizationRequiredError(error)) {
        const response = await client.tools.authorize({
          tool_name: toolName,
          user_id: userId,
        });

        // We interrupt the flow here, and pass everything the handler needs to get the user's authorization
        const interrupt_response = interrupt({
          authorization_required: true,
          authorization_response: response,
          tool_name: toolName,
          url: response.url ?? "",
        });

        // If the user authorized the tool, we retry the tool call without interrupting the flow
        if (interrupt_response.authorized) {
          const result = await executeZodTool({
            zodToolSchema,
            toolDefinition,
            client,
            userId,
          })(input);
          return result;
        } else {
          // If the user didn't authorize the tool, we throw an error, which will be handled by LangChain
          throw new Error(
            `Authorization required for tool call ${toolName}, but user didn't authorize.`
          );
        }
      }
      throw error;
    }
  };
}
````

### Create the tool retrieval helper

The last helper function of this module is the `getTools` helper. This function will take the configurations you defined in the `main.ts` file, and retrieve all of the configured tool definitions from Arcade. Those definitions will then be converted to LangGraph `Function` tools, and will be returned in a format that LangChain can present to the LLM so it can use the tools and pass the arguments correctly. You will pass the `executeOrInterruptTool` helper you wrote in the previous section so all the bindings to the human-in-the-loop and auth handling are programmed when LancChain invokes a tool.


````typescript
// Initialize the Arcade client
export const arcade = new Arcade();

export type GetToolsProps = {
  arcade: Arcade;
  toolkits?: string[];
  tools?: string[];
  userId: string;
  limit?: number;
}


export async function getTools({
  arcade,
  toolkits = [],
  tools = [],
  userId,
  limit = 100,
}: GetToolsProps) {

  if (toolkits.length === 0 && tools.length === 0) {
      throw new Error("At least one tool or toolkit must be provided");
  }

  // Todo(Mateo): Add pagination support
  const from_toolkits = await Promise.all(toolkits.map(async (tkitName) => {
      const definitions = await arcade.tools.list({
          toolkit: tkitName,
          limit: limit
      });
      return definitions.items;
  }));

  const from_tools = await Promise.all(tools.map(async (toolName) => {
      return await arcade.tools.get(toolName);
  }));

  const all_tools = [...from_toolkits.flat(), ...from_tools];
  const unique_tools = Array.from(
      new Map(all_tools.map(tool => [tool.qualified_name, tool])).values()
  );

  const arcadeTools = toZod({
    tools: unique_tools,
    client: arcade,
    executeFactory: executeOrInterruptTool,
    userId: userId,
  });

  // Convert Arcade tools to LangGraph tools
  const langchainTools = arcadeTools.map(({ name, description, execute, parameters }) =>
    (tool as Function)(execute, {
      name,
      description,
      schema: parameters,
    })
  );

  return langchainTools;
}
````

## Building the Agent

Back on the `main.ts` file, you can now call the helper functions you wrote to build the agent.

### Retrieve the configured tools

Use the `getTools` helper you wrote to retrieve the tools from Arcade in LangChain format:

````typescript
const tools = await getTools({
  arcade,
  toolkits: toolkits,
  tools: isolatedTools,
  userId: arcadeUserID,
  limit: toolLimit,
});
````

### Write an interrupt handler

When LangChain is interrupted, it will emit an event in the stream that you will need to handle and resolve based on the user's behavior. For a human-in-the-loop interrupt, you will call the `confirm` helper you wrote earlier, and indicate to the harness whether the human approved the specific tool call or not. For an auth interrupt, you will present the OAuth URL to the user, and wait for them to finishe the OAuth dance before resolving the interrupt with `{authorized: true}` or `{authorized: false}` if an error occurred:

````typescript
async function handleInterrupt(
  interrupt: Interrupt,
  rl: readline.Interface
): Promise<{ authorized: boolean }> {
  const value = interrupt.value;
  const authorization_required = value.authorization_required;
  const hitl_required = value.hitl_required;
  if (authorization_required) {
    const tool_name = value.tool_name;
    const authorization_response = value.authorization_response;
    console.log("⚙️: Authorization required for tool call", tool_name);
    console.log(
      "⚙️: Please authorize in your browser",
      authorization_response.url
    );
    console.log("⚙️: Waiting for you to complete authorization...");
    try {
      await arcade.auth.waitForCompletion(authorization_response.id);
      console.log("⚙️: Authorization granted. Resuming execution...");
      return { authorized: true };
    } catch (error) {
      console.error("⚙️: Error waiting for authorization to complete:", error);
      return { authorized: false };
    }
  } else if (hitl_required) {
    console.log("⚙️: Human in the loop required for tool call", value.tool_name);
    console.log("⚙️: Please approve the tool call", value.input);
    const approved = await confirm("Do you approve this tool call?", rl);
    return { authorized: approved };
  }
  return { authorized: false };
}
````

### Create an Agent instance

Here you create the agent using the `createAgent` function. You pass the system prompt, the model, the tools, and the checkpointer. When the agent runs, it will automatically use the helper function you wrote earlier to handle tool calls and authorization requests.

````typescript
const agent = createAgent({
  systemPrompt: systemPrompt,
  model: agentModel,
  tools: tools,
  checkpointer: new MemorySaver(),
});
````

### Write the invoke helper

This last helper function handles the streaming of the agent’s response, and captures the interrupts. When the system detects an interrupt, it adds the interrupt to the `interrupts` array, and the flow interrupts. If there are no interrupts, it will just stream the agent’s to your console.

````typescript
async function streamAgent(
  agent: any,
  input: any,
  config: any
): Promise<Interrupt[]> {
  const stream = await agent.stream(input, {
    ...config,
    streamMode: "updates",
  });
  const interrupts: Interrupt[] = [];

  for await (const chunk of stream) {
    if (chunk.__interrupt__) {
      interrupts.push(...(chunk.__interrupt__ as Interrupt[]));
      continue;
    }
    for (const update of Object.values(chunk)) {
      for (const msg of (update as any)?.messages ?? []) {
        console.log("🤖: ", msg.toFormattedString());
      }
    }
  }

  return interrupts;
}
````

### Write the main function

Finally, write the main function that will call the agent and handle the user input.

Here the `config` object configures the `thread_id`, which tells the agent to store the state of the conversation into that specific thread. Like any typical agent loop, you:

1. Capture the user input
2. Stream the agent's response
3. Handle any authorization interrupts
4. Resume the agent after authorization
5. Handle any errors
6. Exit the loop if the user wants to quit

````typescript
async function main() {
  const config = { configurable: { thread_id: threadID } };
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  console.log(chalk.green("Welcome to the chatbot! Type 'exit' to quit."));
  while (true) {
    const input = await rl.question("> ");
    if (input.toLowerCase() === "exit") {
      break;
    }
    rl.pause();

    try {
      let agentInput: any = {
        messages: [{ role: "user", content: input }],
      };

      // Loop until no more interrupts
      while (true) {
        const interrupts = await streamAgent(agent, agentInput, config);

        if (interrupts.length === 0) {
          break; // No more interrupts, we're done
        }

        // Handle all interrupts
        const decisions: any[] = [];
        for (const interrupt of interrupts) {
          decisions.push(await handleInterrupt(interrupt, rl));
        }

        // Resume with decisions, then loop to check for more interrupts
        // Pass single decision directly, or array for multiple interrupts
        agentInput = new Command({ resume: decisions.length === 1 ? decisions[0] : decisions });
      }
    } catch (error) {
      console.error(error);
    }

    rl.resume();
  }
  console.log(chalk.red("👋 Bye..."));
  process.exit(0);
}

// Run the main function
main().catch((err) => console.error(err));
````

## Running the Agent

### Run the agent

```bash
bun run main.ts
```

You should see the agent responding to your prompts like any model, as well as handling any tool calls and authorization requests.

## Next Steps

- Clone the [repository](https://github.com/arcade-agents/ts-langchain-Sharepoint) and run it
- Add more toolkits to the `toolkits` array to expand capabilities
- Customize the `systemPrompt` to specialize the agent's behavior
- Explore the [Arcade documentation](https://docs.arcade.dev) for available toolkits

