import os

# Get keys for your project from the project settings page: https://cloud.langfuse.com
os.environ["LANGFUSE_PUBLIC_KEY"] = "pk-lf-101873bd-8396-43e8-841b-c06d20166369" 
os.environ["LANGFUSE_SECRET_KEY"] = "sk-lf-8e5e82b7-b108-4be9-b279-20bec1fd5dbd" 
os.environ["LANGFUSE_HOST"] = "https://cloud.langfuse.com" # 🇪🇺 EU region
# os.environ["LANGFUSE_HOST"] = "https://us.cloud.langfuse.com" # 🇺🇸 US region

from langfuse import get_client
 
langfuse = get_client()
 
# Verify connection
if langfuse.auth_check():
    print("Langfuse client is authenticated and ready!")
else:
    print("Authentication failed. Please check your credentials and host.")


from openinference.instrumentation.smolagents import SmolagentsInstrumentor

SmolagentsInstrumentor().instrument()

from smolagents import CodeAgent, InferenceClientModel

agent = CodeAgent(tools=[], model=InferenceClientModel())
alfred_agent = agent.from_hub('sergiopaniego/AlfredAgent', trust_remote_code=True)
alfred_agent.run("Give me the best playlist for a party at Wayne's mansion. The party idea is a 'villain masquerade' theme")  