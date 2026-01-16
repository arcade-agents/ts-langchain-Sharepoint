from arcadepy import AsyncArcade
from dotenv import load_dotenv
from google.adk import Agent, Runner
from google.adk.artifacts import InMemoryArtifactService
from google.adk.models.lite_llm import LiteLlm
from google.adk.sessions import InMemorySessionService, Session
from google_adk_arcade.tools import get_arcade_tools
from google.genai import types
from human_in_the_loop import auth_tool, confirm_tool_usage

import os

load_dotenv(override=True)


async def main():
    app_name = "my_agent"
    user_id = os.getenv("ARCADE_USER_ID")

    session_service = InMemorySessionService()
    artifact_service = InMemoryArtifactService()
    client = AsyncArcade()

    agent_tools = await get_arcade_tools(
        client, toolkits=["Sharepoint"]
    )

    for tool in agent_tools:
        await auth_tool(client, tool_name=tool.name, user_id=user_id)

    agent = Agent(
        model=LiteLlm(model=f"openai/{os.environ["OPENAI_MODEL"]}"),
        name="google_agent",
        instruction="# Introduction
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
        description="An agent that uses Sharepoint tools provided to perform any task",
        tools=agent_tools,
        before_tool_callback=[confirm_tool_usage],
    )

    session = await session_service.create_session(
        app_name=app_name, user_id=user_id, state={
            "user_id": user_id,
        }
    )
    runner = Runner(
        app_name=app_name,
        agent=agent,
        artifact_service=artifact_service,
        session_service=session_service,
    )

    async def run_prompt(session: Session, new_message: str):
        content = types.Content(
            role='user', parts=[types.Part.from_text(text=new_message)]
        )
        async for event in runner.run_async(
            user_id=user_id,
            session_id=session.id,
            new_message=content,
        ):
            if event.content.parts and event.content.parts[0].text:
                print(f'** {event.author}: {event.content.parts[0].text}')

    while True:
        user_input = input("User: ")
        if user_input.lower() == "exit":
            print("Goodbye!")
            break
        await run_prompt(session, user_input)


if __name__ == '__main__':
    import asyncio
    asyncio.run(main())