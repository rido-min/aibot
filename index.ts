import { App } from '@microsoft/teams.apps'
import { openai } from '@ai-sdk/openai'
import { streamText } from 'ai'

const app = new App()

app.on('message', async ({ send, stream, activity }) => {
  if (!activity.text) {
    await send('Please send a text message.')
    return
  }

  stream.update('Thinking...')

  try {
    const result = streamText({
      model: openai('gpt-4o'),
      messages: [{ role: 'user', content: activity.text }],
    })

    for await (const chunk of result.textStream) {
      stream.emit(chunk)
    }

    await stream.close()
  } catch (err) {
    await send(`Sorry, something went wrong: ${err instanceof Error ? err.message : String(err)}`)
  }
})

app.start()
