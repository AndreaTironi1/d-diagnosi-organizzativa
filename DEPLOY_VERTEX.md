# Deploy to Google Cloud (Vertex AI)

Questa guida spiega come pubblicare l'applicazione su Google Cloud usando Cloud Run.

## Prerequisiti

1. Account Google Cloud attivo
2. [Google Cloud CLI (gcloud)](https://cloud.google.com/sdk/docs/install) installato
3. Progetto Google Cloud creato
4. Fatturazione abilitata sul progetto

## Opzione 1: Deploy con Cloud Run (Consigliato)

Cloud Run è un servizio serverless che scala automaticamente ed è perfetto per questa applicazione.

### Passo 1: Configurazione iniziale

```bash
# Login a Google Cloud
gcloud auth login

# Imposta il progetto
gcloud config set project YOUR_PROJECT_ID

# Abilita le API necessarie
gcloud services enable run.googleapis.com
gcloud services enable containerregistry.googleapis.com
gcloud services enable cloudbuild.googleapis.com
```

### Passo 2: Build e Deploy

```bash
# Naviga alla directory del progetto
cd C:\Users\a.tironi\Downloads\ClaudeCode

# Deploy diretto (Cloud Build farà il build automaticamente)
gcloud run deploy prompt-executor \
  --source . \
  --region europe-west1 \
  --platform managed \
  --allow-unauthenticated \
  --set-env-vars ANTHROPIC_API_KEY=your_api_key_here \
  --memory 512Mi \
  --timeout 300
```

### Passo 3: Configurare variabili d'ambiente

Dopo il deploy, puoi aggiornare le variabili d'ambiente:

```bash
gcloud run services update prompt-executor \
  --region europe-west1 \
  --update-env-vars ANTHROPIC_API_KEY=your_actual_api_key
```

**IMPORTANTE**: Non includere mai la tua API key nel codice o nel Dockerfile!

### Opzione alternativa: Usare Secret Manager (Più Sicuro)

```bash
# Abilita Secret Manager API
gcloud services enable secretmanager.googleapis.com

# Crea un secret per la API key
echo -n "your_api_key_here" | gcloud secrets create anthropic-api-key --data-file=-

# Deploy con secret
gcloud run deploy prompt-executor \
  --source . \
  --region europe-west1 \
  --platform managed \
  --allow-unauthenticated \
  --update-secrets ANTHROPIC_API_KEY=anthropic-api-key:latest \
  --memory 512Mi \
  --timeout 300
```

## Opzione 2: Deploy manuale con Docker

### Passo 1: Build dell'immagine Docker

```bash
# Build locale
docker build -t gcr.io/YOUR_PROJECT_ID/prompt-executor .

# Testa localmente
docker run -p 8080:8080 -e ANTHROPIC_API_KEY=your_api_key gcr.io/YOUR_PROJECT_ID/prompt-executor
```

### Passo 2: Push a Container Registry

```bash
# Configura Docker per Google Cloud
gcloud auth configure-docker

# Push dell'immagine
docker push gcr.io/YOUR_PROJECT_ID/prompt-executor
```

### Passo 3: Deploy a Cloud Run

```bash
gcloud run deploy prompt-executor \
  --image gcr.io/YOUR_PROJECT_ID/prompt-executor \
  --region europe-west1 \
  --platform managed \
  --allow-unauthenticated \
  --set-env-vars ANTHROPIC_API_KEY=your_api_key_here \
  --memory 512Mi \
  --timeout 300
```

## Opzione 3: Deploy con Artifact Registry (Raccomandato per produzione)

```bash
# Abilita Artifact Registry
gcloud services enable artifactregistry.googleapis.com

# Crea repository
gcloud artifacts repositories create prompt-executor-repo \
  --repository-format=docker \
  --location=europe-west1

# Configura autenticazione
gcloud auth configure-docker europe-west1-docker.pkg.dev

# Build e push
docker build -t europe-west1-docker.pkg.dev/YOUR_PROJECT_ID/prompt-executor-repo/prompt-executor:v1 .
docker push europe-west1-docker.pkg.dev/YOUR_PROJECT_ID/prompt-executor-repo/prompt-executor:v1

# Deploy
gcloud run deploy prompt-executor \
  --image europe-west1-docker.pkg.dev/YOUR_PROJECT_ID/prompt-executor-repo/prompt-executor:v1 \
  --region europe-west1 \
  --platform managed \
  --allow-unauthenticated \
  --memory 512Mi \
  --timeout 300
```

## Configurazione Avanzata

### Aumentare limiti per file Excel grandi

```bash
gcloud run services update prompt-executor \
  --region europe-west1 \
  --memory 1Gi \
  --cpu 2 \
  --timeout 600 \
  --max-instances 10
```

### Limitare accesso (autenticazione)

```bash
# Rimuovi accesso pubblico
gcloud run services remove-iam-policy-binding prompt-executor \
  --region europe-west1 \
  --member="allUsers" \
  --role="roles/run.invoker"

# Aggiungi utenti specifici
gcloud run services add-iam-policy-binding prompt-executor \
  --region europe-west1 \
  --member="user:email@example.com" \
  --role="roles/run.invoker"
```

### Configurare dominio personalizzato

```bash
# Mappa un dominio personalizzato
gcloud run domain-mappings create \
  --service prompt-executor \
  --domain your-domain.com \
  --region europe-west1
```

## Monitoraggio e Logging

### Visualizzare i logs

```bash
# Logs in tempo reale
gcloud run services logs tail prompt-executor --region europe-west1

# Logs recenti
gcloud run services logs read prompt-executor --region europe-west1 --limit 50
```

### Metriche

Visualizza metriche su Cloud Console:
```
https://console.cloud.google.com/run/detail/europe-west1/prompt-executor/metrics
```

## Costi Stimati

Cloud Run pricing (pay-per-use):
- **Request**: $0.40 per million requests
- **CPU**: $0.00002400 per vCPU-second
- **Memory**: $0.00000250 per GiB-second
- **Free tier**: 2 milioni di richieste/mese

Per questa applicazione, il costo sarà minimo se usata saltuariamente.

## Variabili d'ambiente disponibili

- `ANTHROPIC_API_KEY`: API key per Claude (obbligatorio)
- `PORT`: Porta del server (Cloud Run usa automaticamente 8080)

## Troubleshooting

**Errore "Container failed to start":**
- Verifica che PORT sia impostato correttamente (8080 per Cloud Run)
- Controlla i logs: `gcloud run services logs read prompt-executor --region europe-west1`

**Timeout durante batch processing:**
- Aumenta il timeout: `--timeout 600` (max 60 minuti per 2nd gen)
- Considera l'uso di Cloud Run 2nd generation

**Out of memory:**
- Aumenta la memoria: `--memory 2Gi`

**ANTHROPIC_API_KEY non trovata:**
- Verifica che la variabile sia impostata correttamente
- Se usi Secret Manager, verifica i permessi

## Sicurezza

1. **Non committare mai .env nel repository**
2. **Usa Secret Manager per le API keys**
3. **Abilita Cloud Armor per protezione DDoS**
4. **Configura autenticazione IAM per accesso limitato**
5. **Abilita HTTPS (automatico su Cloud Run)**

## Aggiornamenti

Per aggiornare l'applicazione:

```bash
# Metodo 1: Deploy diretto
gcloud run deploy prompt-executor --source . --region europe-west1

# Metodo 2: Con nuova immagine Docker
docker build -t europe-west1-docker.pkg.dev/YOUR_PROJECT_ID/prompt-executor-repo/prompt-executor:v2 .
docker push europe-west1-docker.pkg.dev/YOUR_PROJECT_ID/prompt-executor-repo/prompt-executor:v2
gcloud run services update prompt-executor \
  --image europe-west1-docker.pkg.dev/YOUR_PROJECT_ID/prompt-executor-repo/prompt-executor:v2 \
  --region europe-west1
```

## Rollback

Se qualcosa va storto:

```bash
# Lista le revisioni
gcloud run revisions list --service prompt-executor --region europe-west1

# Rollback a revisione precedente
gcloud run services update-traffic prompt-executor \
  --to-revisions REVISION_NAME=100 \
  --region europe-west1
```

## URL dell'applicazione

Dopo il deploy, riceverai un URL tipo:
```
https://prompt-executor-[hash]-ew.a.run.app
```

## Supporto

- [Cloud Run Documentation](https://cloud.google.com/run/docs)
- [Google Cloud Support](https://cloud.google.com/support)
- [Cloud Run Pricing](https://cloud.google.com/run/pricing)
