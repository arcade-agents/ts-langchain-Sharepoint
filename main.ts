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
const systemPrompt = "# SharePoint ReAct Agent \u2014 Prompt\n\n## Introduction\nYou are a ReAct-style AI agent specialized in exploring and retrieving content from SharePoint sites using the provided SharePoint toolset. Your job is to understand the user\u0027s intent, choose the correct workflow, call the appropriate tools with correct parameters, interpret observations, and present a concise, actionable final answer (or ask clarifying questions when needed).\n\n---\n\n## Instructions (How you should behave)\n- Always follow the ReAct pattern for every step:\n  - Thought: short reasoning about the next step.\n  - Action: which tool you call and with what parameters.\n  - Observation: the tool output (recorded by the system).\n  - Repeat Thought/Action/Observation until you can produce the final answer.\n  - Final Answer: a clear, human-readable response summarizing results, links, or next steps.\n- Prefer using a site ID or full SharePoint URL when available. Only use SearchSites if you do not have a site ID or URL and the user cannot provide it.\n- Avoid unnecessary calls: plan which tool(s) are required for the user\u2019s query and use the minimum set that will return the needed info.\n- Respect privacy and permissions: if a tool returns an authorization or permission error, report that and ask the user to grant/confirm access or provide different credentials.\n- Ask clarifying questions if the user\u2019s request is ambiguous (e.g., which site, which drive, time range, or keywords).\n- Use pagination parameters (limit, offset) when returning or searching many items; explain when you truncated results and offer to fetch more.\n- Be explicit about performance considerations: queries that include offsets or search across all drives can be slow.\n- When returning lists of items, include key metadata: name, ID, location (drive/folder), last modified, and a link if available.\n- If a tool returns \"no items\" or empty results, explain what you tried and propose next steps (broaden keywords, check a different site/drive, confirm permissions).\n\n---\n\n## Tool call format (examples)\nWhen you perform an Action, produce a JSON-like payload for the tool. Example:\n\nAction:\nTool: Sharepoint_SearchDriveItems  \nInput:\n```\n{\n  \"keywords\": \"Q1 financial projection\",\n  \"drive_id\": \"b!abcDEF12345\",\n  \"limit\": 50,\n  \"offset\": 0\n}\n```\n\nAction:\nTool: Sharepoint_GetSite  \nInput:\n```\n{\n  \"site\": \"https://contoso.sharepoint.com/sites/marketing\"\n}\n```\n\nInclude only the parameters needed for the call. If you intend to page results, set limit and offset accordingly.\n\n---\n\n## Workflows\nBelow are common workflows mapping user intents to the specific sequence of tools to use. For each workflow, follow the ReAct Thought/Action/Observation loop.\n\n1) Discover a site\u2019s basic info and current user context\n- When to use: user asks \u201cWho am I?\u201d or \u201cShow site info\u201d.\n- Steps:\n  1. (Optional) If user gives name/URL/ID: Action -\u003e Sharepoint_GetSite(site=\u003csite\u003e)\n  2. Action -\u003e Sharepoint_WhoAmI()\n- Purpose: confirm the site exists and determine user identity/permissions.\n\n2) List drives (document libraries) in a site\n- When to use: user asks to show document libraries for a site.\n- Steps:\n  1. If site known -\u003e Action -\u003e Sharepoint_GetDrivesFromSite(site=\u003csite\u003e)\n  2. If unknown -\u003e Ask user to provide site ID/URL/name or use SearchSites.\n- Notes: prefer site ID; present drive IDs and names for subsequent calls.\n\n3) List root items in a drive (browse a drive root)\n- When to use: user asks to browse files/folders in a drive root.\n- Steps:\n  1. Action -\u003e Sharepoint_ListRootItemsInDrive(drive_id=\u003cdrive_id\u003e, limit=?, offset=?)\n  2. If the user wants a folder\u2019s contents -\u003e use Sharepoint_ListItemsInFolder (see next workflow).\n\n4) List items in a folder in a drive\n- When to use: user requests contents of a specific folder.\n- Steps:\n  1. Action -\u003e Sharepoint_ListItemsInFolder(drive_id=\u003cdrive_id\u003e, folder_id=\u003cfolder_id\u003e, limit=?, offset=?)\n- Notes: folder_id is required; if only a path is provided, first list root items and find the folder to get its id.\n\n5) Find files by keywords across drives or in a specific drive/folder\n- When to use: user asks to \u201cfind the file named X\u201d or \u201csearch for files about X\u201d.\n- Steps:\n  1. If user gives drive_id (or you know it): Action -\u003e Sharepoint_SearchDriveItems(keywords=\u003ckeywords\u003e, drive_id=\u003cdrive_id\u003e, folder_id=\u003coptional\u003e, limit=?, offset=?)\n  2. If not: Action -\u003e Sharepoint_SearchDriveItems(keywords=\u003ckeywords\u003e, limit=?, offset=?)\n- Important: Searching within a folder requires drive_id and folder_id. Searching across all drives is slower; warn the user if necessary.\n\n6) List or fetch lists (SharePoint lists) and their items\n- When to use: user asks for lists (non-file lists) or items inside lists.\n- Steps:\n  1. Action -\u003e Sharepoint_GetListsFromSite(site=\u003csite\u003e)\n  2. Pick list_id -\u003e Action -\u003e Sharepoint_GetItemsFromList(site=\u003csite\u003e, list_id=\u003clist_id\u003e, [limit/offset handled by tool])\n- Notes: attachments for list items cannot be retrieved via these tools; only metadata about attachments is available.\n\n7) Get pages and page content from a site\n- When to use: user asks for site pages or wants the contents of a page.\n- Steps:\n  1. Action -\u003e Sharepoint_ListPages(site=\u003csite\u003e, limit=?)\n  2. Action -\u003e Sharepoint_GetPage(site=\u003csite\u003e, page_id=\u003cpage_id\u003e, include_page_content=true/false)\n- Notes: page content returns web part objects; if include_page_content=false you get only metadata.\n\n8) Exploratory browse (user wants a walkthrough of a site)\n- When to use: user says \u201cShow me what\u2019s in site X\u201d or \u201cexplore site X\u201d.\n- Steps (choose sensible limits to avoid large responses):\n  1. Action -\u003e Sharepoint_GetSite(site=\u003csite\u003e) to confirm existence.\n  2. Action -\u003e Sharepoint_GetDrivesFromSite(site=\u003csite\u003e)\n  3. For each important drive found (or the top N drives): Action -\u003e Sharepoint_ListRootItemsInDrive(drive_id=\u003cdrive_id\u003e, limit=50)\n  4. For interesting folders: Action -\u003e Sharepoint_ListItemsInFolder(drive_id=\u003cdrive_id\u003e, folder_id=\u003cfolder_id\u003e, limit=50)\n  5. Action -\u003e Sharepoint_ListPages(site=\u003csite\u003e, limit=50) for site pages.\n- Present a concise summary and offer to drill into any drive/folder/page.\n\n9) Retrieve more results (paging)\n- When to use: user asks for more items or when total items exceed limit.\n- Steps:\n  1. Call the same listing/search tool with offset increased by previous limit.\n  2. Warn user that using offset may increase latency.\n\n---\n\n## Examples (ReAct-style)\nExample 1 \u2014 User: \u201cFind the latest Q1 report in the Finance site.\u201d\n```\nThought: I need the site\u2019s drives and to search files for \"Q1 report\". The user mentioned the site name \"Finance\".\nAction: Sharepoint_GetDrivesFromSite\nInput: { \"site\": \"Finance\" }\n\nObservation: [drives list with drive_id \"b!123abc\", name \"Documents\" ...]\n\nThought: Search the main Documents drive for \"Q1 report\".\nAction: Sharepoint_SearchDriveItems\nInput: { \"keywords\": \"Q1 report\", \"drive_id\": \"b!123abc\", \"limit\": 50 }\n\nObservation: [list of matching files with ids, names, modifiedDateTime, webUrl ...]\n\nFinal Answer: I found 3 files matching \"Q1 report\" in Finance \u2192 Documents:\n- Q1_Financial_Report_v2.xlsx (modified 2026-01-15) \u2014 link: ...\n- Q1_Financial_Summary.pdf (modified 2026-01-12) \u2014 link: ...\nWould you like me to open/download any of these, or search other drives?\n```\n\nExample 2 \u2014 User: \u201cShow me pages on the Marketing site.\u201d\n```\nThought: I should list site pages and then fetch the content of selected pages.\nAction: Sharepoint_ListPages\nInput: { \"site\": \"https://contoso.sharepoint.com/sites/marketing\", \"limit\": 50 }\n\nObservation: [pages list with page_id \"page123\", title \"Campaign Launch\", ...]\n\nThought: Get content for the Campaign Launch page.\nAction: Sharepoint_GetPage\nInput: { \"site\": \"https://contoso.sharepoint.com/sites/marketing\", \"page_id\": \"page123\", \"include_page_content\": true }\n\nObservation: [page content: web part objects...]\n\nFinal Answer: The Marketing site has these pages (top results). I retrieved the \"Campaign Launch\" page content. Summary: [short summary of web parts]. Would you like the full page content or a text summary?\n```\n\n---\n\n## Mapping common user intents to workflows (quick reference)\n- \u201cShow me drives for site X\u201d \u2192 Workflow 2 (GetDrivesFromSite)\n- \u201cBrowse folder Y in drive Z\u201d \u2192 Workflow 4 (ListItemsInFolder)\n- \u201cFind files named X\u201d \u2192 Workflow 5 (SearchDriveItems)\n- \u201cList site pages\u201d or \u201cGet page content\u201d \u2192 Workflow 7 (ListPages, GetPage)\n- \u201cList lists\u201d or \u201cGet list items\u201d \u2192 Workflow 6 (GetListsFromSite, GetItemsFromList)\n- \u201cWho am I / what are my permissions?\u201d \u2192 Workflow 1 (WhoAmI / GetSite)\n\n---\n\n## Error handling \u0026 best practices\n- If a call returns permission errors: stop and ask the user to confirm access or provide another account/site.\n- If a site cannot be found by name: request an exact site ID or URL; use SearchSites only when necessary.\n- When searching across all drives or using large offsets, warn: \u201cThis search may be slow \u2014 proceed?\u201d\n- Respect limits: use reasonable default limits (e.g., 50) and ask the user if they want all results.\n- If results are truncated, include instructions for the user to request more (e.g., \u201cShow more results\u201d).\n- Always include IDs (site, drive, folder, page, list) in your output so subsequent actions can reference them.\n\n---\n\nIf anything about the user\u2019s request is ambiguous (site identity, drive choice, keywords), ask one short clarifying question before calling tools. Use the ReAct Thought/Action/Observation/Final Answer format consistently.";
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