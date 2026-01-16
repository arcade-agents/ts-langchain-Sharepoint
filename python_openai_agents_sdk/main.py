from agents import (Agent, Runner, AgentHooks, Tool, RunContextWrapper,
                    TResponseInputItem,)
from functools import partial
from arcadepy import AsyncArcade
from agents_arcade import get_arcade_tools
from typing import Any
from human_in_the_loop import (UserDeniedToolCall,
                               confirm_tool_usage,
                               auth_tool)

import globals


class CustomAgentHooks(AgentHooks):
    def __init__(self, display_name: str):
        self.event_counter = 0
        self.display_name = display_name

    async def on_start(self,
                       context: RunContextWrapper,
                       agent: Agent) -> None:
        self.event_counter += 1
        print(f"### ({self.display_name}) {
              self.event_counter}: Agent {agent.name} started")

    async def on_end(self,
                     context: RunContextWrapper,
                     agent: Agent,
                     output: Any) -> None:
        self.event_counter += 1
        print(
            f"### ({self.display_name}) {self.event_counter}: Agent {
                # agent.name} ended with output {output}"
                agent.name} ended"
        )

    async def on_handoff(self,
                         context: RunContextWrapper,
                         agent: Agent,
                         source: Agent) -> None:
        self.event_counter += 1
        print(
            f"### ({self.display_name}) {self.event_counter}: Agent {
                source.name} handed off to {agent.name}"
        )

    async def on_tool_start(self,
                            context: RunContextWrapper,
                            agent: Agent,
                            tool: Tool) -> None:
        self.event_counter += 1
        print(
            f"### ({self.display_name}) {self.event_counter}:"
            f" Agent {agent.name} started tool {tool.name}"
            f" with context: {context.context}"
        )

    async def on_tool_end(self,
                          context: RunContextWrapper,
                          agent: Agent,
                          tool: Tool,
                          result: str) -> None:
        self.event_counter += 1
        print(
            f"### ({self.display_name}) {self.event_counter}: Agent {
                # agent.name} ended tool {tool.name} with result {result}"
                agent.name} ended tool {tool.name}"
        )


async def main():

    context = {
        "user_id": os.getenv("ARCADE_USER_ID"),
    }

    client = AsyncArcade()

    arcade_tools = await get_arcade_tools(
        client, toolkits=["Sharepoint"]
    )

    for tool in arcade_tools:
        # - human in the loop
        if tool.name in ENFORCE_HUMAN_CONFIRMATION:
            tool.on_invoke_tool = partial(
                confirm_tool_usage,
                tool_name=tool.name,
                callback=tool.on_invoke_tool,
            )
        # - auth
        await auth_tool(client, tool.name, user_id=context["user_id"])

    agent = Agent(
        name="",
        instructions="# Introduction
Welcome to the SharePoint AI Agent! This agent is designed to help you interact with SharePoint environments effectively. Whether you need to retrieve sites, lists, items, or pages, this agent can streamline your workflow and provide you with the information you need quickly and efficiently.

# Instructions
1. **Understand the User's Needs**: Begin by identifying what the user wants to achieve with SharePoint. This could include accessing sites, fetching lists, retrieving items, etc.
2. **Select Appropriate Tools**: Based on the user’s request, choose the relevant SharePoint tools to fulfill the requirements.
3. **Execute Workflows**: Follow the designated workflows to gather information and respond to the user’s queries.
4. **Provide Clear Outputs**: Summarize the results in a user-friendly manner, ensuring clarity and usability of the provided information.

# Workflows

## Workflow 1: Retrieve User Information
1. Use the **Sharepoint_WhoAmI** tool to get information about the current user and their SharePoint environment.

## Workflow 2: List SharePoint Sites
1. Use the **Sharepoint_ListSites** tool to list all SharePoint sites accessible to the current user.
    - Parameters: (optional) `limit`, `offset`

## Workflow 3: Get Drives from a Specific Site
1. Use the **Sharepoint_GetDrivesFromSite** tool with the site name or ID to retrieve document libraries from the specified SharePoint site.

## Workflow 4: Retrieve Lists from a Site
1. Use the **Sharepoint_GetListsFromSite** tool to get all lists available in a specified SharePoint site.

## Workflow 5: Retrieve Items from a List
1. Use the **Sharepoint_GetItemsFromList** tool to retrieve items from a specific list.
    - Parameters: `site`, `list_id`

## Workflow 6: Search Items in Drives
1. Use the **Sharepoint_SearchDriveItems** tool to search for specific items in one or more SharePoint drives using keywords.
    - Parameters: `keywords`, (optional) `drive_id`, `folder_id`, `limit`, `offset`

## Workflow 7: Get Page Content or Metadata
1. Use the **Sharepoint_GetPage** tool to retrieve metadata and page content from a specific page in a given SharePoint site.
    - Parameters: `site`, `page_id`, (optional) `include_page_content`

## Workflow 8: Retrieve Items from a Folder or Drive
1. Use the **Sharepoint_ListItemsInFolder** tool to get items from a specific folder within a drive.
    - Parameters: `drive_id`, `folder_id`, (optional) `limit`, `offset`
2. Alternatively, use **Sharepoint_ListRootItemsInDrive** to retrieve items from the root of a specified drive.

Use these workflows and tools iteratively based on the user's requests to ensure a seamless SharePoint experience!",
        model=os.environ["OPENAI_MODEL"],
        tools=arcade_tools,
        hooks=CustomAgentHooks(display_name="")
    )

    # initialize the conversation
    history: list[TResponseInputItem] = []
    # run the loop!
    while True:
        prompt = input("You: ")
        if prompt.lower() == "exit":
            break
        history.append({"role": "user", "content": prompt})
        try:
            result = await Runner.run(
                starting_agent=agent,
                input=history,
                context=context
            )
            history = result.to_input_list()
            print(result.final_output)
        except UserDeniedToolCall as e:
            history.extend([
                {"role": "assistant",
                 "content": f"Please confirm the call to {e.tool_name}"},
                {"role": "user",
                 "content": "I changed my mind, please don't do it!"},
                {"role": "assistant",
                 "content": f"Sure, I cancelled the call to {e.tool_name}."
                 " What else can I do for you today?"
                 },
            ])
            print(history[-1]["content"])

if __name__ == "__main__":
    import asyncio

    asyncio.run(main())