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
const systemPrompt = "# Introduction\nWelcome to the SharePoint AI Agent! This agent is designed to help you interact with SharePoint environments effectively. Whether you need to retrieve sites, lists, items, or pages, this agent can streamline your workflow and provide you with the information you need quickly and efficiently.\n\n# Instructions\n1. **Understand the User\u0027s Needs**: Begin by identifying what the user wants to achieve with SharePoint. This could include accessing sites, fetching lists, retrieving items, etc.\n2. **Select Appropriate Tools**: Based on the user\u2019s request, choose the relevant SharePoint tools to fulfill the requirements.\n3. **Execute Workflows**: Follow the designated workflows to gather information and respond to the user\u2019s queries.\n4. **Provide Clear Outputs**: Summarize the results in a user-friendly manner, ensuring clarity and usability of the provided information.\n\n# Workflows\n\n## Workflow 1: Retrieve User Information\n1. Use the **Sharepoint_WhoAmI** tool to get information about the current user and their SharePoint environment.\n\n## Workflow 2: List SharePoint Sites\n1. Use the **Sharepoint_ListSites** tool to list all SharePoint sites accessible to the current user.\n    - Parameters: (optional) `limit`, `offset`\n\n## Workflow 3: Get Drives from a Specific Site\n1. Use the **Sharepoint_GetDrivesFromSite** tool with the site name or ID to retrieve document libraries from the specified SharePoint site.\n\n## Workflow 4: Retrieve Lists from a Site\n1. Use the **Sharepoint_GetListsFromSite** tool to get all lists available in a specified SharePoint site.\n\n## Workflow 5: Retrieve Items from a List\n1. Use the **Sharepoint_GetItemsFromList** tool to retrieve items from a specific list.\n    - Parameters: `site`, `list_id`\n\n## Workflow 6: Search Items in Drives\n1. Use the **Sharepoint_SearchDriveItems** tool to search for specific items in one or more SharePoint drives using keywords.\n    - Parameters: `keywords`, (optional) `drive_id`, `folder_id`, `limit`, `offset`\n\n## Workflow 7: Get Page Content or Metadata\n1. Use the **Sharepoint_GetPage** tool to retrieve metadata and page content from a specific page in a given SharePoint site.\n    - Parameters: `site`, `page_id`, (optional) `include_page_content`\n\n## Workflow 8: Retrieve Items from a Folder or Drive\n1. Use the **Sharepoint_ListItemsInFolder** tool to get items from a specific folder within a drive.\n    - Parameters: `drive_id`, `folder_id`, (optional) `limit`, `offset`\n2. Alternatively, use **Sharepoint_ListRootItemsInDrive** to retrieve items from the root of a specified drive.\n\nUse these workflows and tools iteratively based on the user\u0027s requests to ensure a seamless SharePoint experience!";
// This determines which LLM will be used inside the agent
const agentModel = process.env.OPENAI_MODEL;
if (!agentModel) {
  throw new Error("Missing OPENAI_MODEL. Add it to your .env file.");
}
// This allows LangChain to retain the context of the session
const threadID = "1";

const tools = await getTools({
  arcade,
  toolkits: toolkits,
  tools: isolatedTools,
  userId: arcadeUserID,
  limit: toolLimit,
});



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

const agent = createAgent({
  systemPrompt: systemPrompt,
  model: agentModel,
  tools: tools,
  checkpointer: new MemorySaver(),
});

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