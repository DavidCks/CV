# CADDi Interview 1

## Motivation

### Freelancing

#### Positives (Freelancing)

- lucrative

#### Negatives (Freelancing)

- unstable
- not technically challenging
- lots of non-engineering overhead
  - marketing
  - prospecting
  - sales
  - client management
  - accounting
  - etc.

#### Want

- stable workplace
- stable income
- focus on engineering
  - solving complex problems
  - turning impactful visions into reality

## Values

- having an impact
- growing as a developer
- building together with an ambitious team

### Sidenote

#### Solo work

##### Positives (Solo work)

- fast
- efficient

##### Positives (Team work)

- better results
- better products
  - bigger impact

## Other companies

- in talks with multiple companies

## CADDi (compared to other companies)

- ambitious vision
- growth opportunity

## Timeline

- Within the next 2 months (available as soon as possible)

## Relevant experience

### CADDi needs - leadership, initiative, and apis

- take the lead from the perspective of the infrastructure in building a data collection pipeline
- promoting data utilization
- building APIs

#### Leadership and initiative

- Took lead and initiative in building services using AWS
  - Particularly, built an API for generating TTS for Palestinian Arabic (PATTS)
    - python3
    - FastAPI
    - AWS Polly
    - Supabase
    - Docker

Whitepaper: <https://cjki.org/download/download/patts.html>

##### Data lifecycle

1. Users input palestinian arabic text
2. Text gets converted to IPA (using AI model in the future)
   - Currently using a rule-based system
   - Rule-based system is not perfect
   - AI model is not perfect either
   - Need to collect data to improve the AI model
3. IPA gets converted to audio
4. Audio gets stored in cloud storage (S3-like)
5. Users can download and evaluate audio / correct IPA
6. IPA gets stored in a database
7. AI model gets improved based on user feedback

##### Sidenote (PATTS)

- Certain audio and IPA data made available to the public through application
  - Palestinian Arabic Verb Conjugator (PAVE)
    - Dart
    - Flutter
    - Docker
    - Supabase
    - nginx

- Building APIs and batch infrastructure for using machine learning in systems

### CADDi needs - batch processing, CI/CD and performance tuning

#### Batch processing

- Built multi-threaded batch processing system for generating text data
  - python3
  - OpenAI API

#### CI/CD

- Build various CI/CD pipelines
  - Auto-deployment of changes in CMS system on customers website
  - Auto-deployment of websites on code changes by default for all customers
  - Automated unit tests on Commits/PRs for rust packages

#### Performance tuning

##### On-device performance tuning

- Performance is very important when building mobile applications
  - Mobile devices have limited resources
  - Mobile devices range from low-end to high-end
    - Performance tuning is crucial for a good user experience on low-end devices

- Built an app framework with automated performance tuning by default
  - on-device performance measurements to optimize application parameters at runtime
    - Particularly for optimizing search algorithms
    - adjusting search limits in SQLite based on device performance

(Dart, Flutter, SQLite)

##### Database performance tuning

- Analyzing search algorithms and building indices based on search patterns, particularly:
  - analyzing the data structures that guide the behavour of the search algorithms
  - automatically creating indices based on the data structures

### CADDi needs - ML processing pipelines, cost optimization, process documentation

#### ML processing pipelines

- Built a pipeline for building datasets for TTS model training (for piperTTS)
  - Input: 30s audio
  - Output: 4h audio dataset in LJSpeech format

- Built a pipeline for 3D AI Characters with locally AI generated
  - Responses (using Llama3.2) (WIP)
  - Audio
  - Emotion inference
  - Animation selection
  - Animation generation (using momask) (WIP)

#### Cost optimization

- All of my projects are cost optimized
  - Everything on-device that can comfortably run on-device
  - Effectively using cloud resources and pricing plans
    - Utilizing free tiers (AWS, Supabase, etc.)
    - Avoiding redundancy, for example
      - Cutting audio files into chunks and accessing them from a database rather than regenerating them
      - Using open-source tools and self-hosting them
        - For example, using local LLMs rather than using OpenAI API
        - Using self-hosted TTS rather than using AWS Polly
        - Self-hosting Supabase rather than using Supabase cloud

#### Process documentation

- Detailed documentation is my default
  - All of my public packages are documented
  - Where applicable, I create detailed schematics to ensure efficient communication, for example
    - For the LRAG project, I created a detailed schematic of the process flow <https://www.cjk.org/wp-content/uploads/LRAG_Japanese.pdf>
    - For the data engineers at CJKI, I created detailed and easy to understand requrests for data changes
      - Defining data structures visually
      - Defining both precise changes and the reasoning behind them (for context)

##### Sidenote (Process documentation and communication)

- I am a strong believer in process documentation and communication
- Thanks to my broad experience, I have a good understanding of the different roles in a team and their needs, particularly:
  - Designers
  - Frontend developers
  - Backend developers
  - Data engineers

### Fun sidenote

- I automate everything
  - voice-activated my pc when I was 17 (when amazon alexa first came out), through which I learned about:
    - speech recognition
    - natural language processing
    - server management
    - DNS management
    - SSL certificates
    - php extensions
  - automated home appliances for prointernet at 18
    - lighting
    - watering systems
    - heating
    - etc.
  - automated creation of various documents from a single data source, most recently my
    - CV
    - Resume

### Questions

You use WASM and WebGL? Big fan!

- I use WASM and WebGL for various projects, such as
  - Running AI models in the browser
  - Building 3D applications

What does CADDi use WASM and WebGL for?

---

What will be my responsibilities when I get hired?

### Other comments

- I am a big fan of CADDi's vision and mission, particularly improving work efficiency
  - App framework creation for CJKI was done as my own initiative
  - Development time for new applications went from months to weeks

The development environment matches my skillset to a very high degree. I am excited to get to apply my broad skillset in a professional environment and to have an impact doing doing what I love:

- automating stuff,
- making stuff more efficient, and
- having a big impact on technology that will change the world
