import express from 'express';
import { BailianClient } from './bailian-client.js';

const config = {
  accessKeyId: process.env.ALIBABA_CLOUD_ACCESS_KEY_ID || '',
  accessKeySecret: process.env.ALIBABA_CLOUD_ACCESS_KEY_SECRET || '',
  endpoint: process.env.BAILIAN_ENDPOINT || 'bailian.cn-beijing.aliyuncs.com',
  workspaceId: process.env.BAILIAN_WORKSPACE_ID || '',
};

const missing: string[] = [];
if (!config.accessKeyId) missing.push('ALIBABA_CLOUD_ACCESS_KEY_ID');
if (!config.accessKeySecret) missing.push('ALIBABA_CLOUD_ACCESS_KEY_SECRET');
if (!config.workspaceId) missing.push('BAILIAN_WORKSPACE_ID');
if (missing.length > 0) {
  console.error(`FATAL: Missing required environment variables: ${missing.join(', ')}`);
  process.exit(1);
}

const client = new BailianClient(config);
const app = express();
const port = parseInt(process.env.PORT || '3000', 10);

app.get('/healthz', (_req, res) => {
  res.json({ status: 'ok', service: 'bailian-knowledge-base-http' });
});

// -------- 工具 1: 列出知识库 --------
app.get('/list_knowledge_bases', async (req, res) => {
  try {
    const params: Record<string, unknown> = {};
    if (req.query.index_name) params.IndexName = String(req.query.index_name);
    if (req.query.page_number) params.PageNumber = Number(req.query.page_number);
    if (req.query.page_size) params.PageSize = Number(req.query.page_size);

    const result = await client.listIndices(params as any);
    res.json(result);
  } catch (error) {
    res.status(500).json({ error: (error as Error).message });
  }
});

// -------- 工具 2: 检索知识库 --------
app.post('/retrieve_knowledge', async (req, res) => {
  try {
    const body = req.body as Record<string, unknown>;
    if (!body.index_id || !body.query) {
      res.status(400).json({ error: 'index_id and query are required' });
      return;
    }

    const params: Record<string, unknown> = {
      IndexId: body.index_id as string,
      Query: body.query as string,
    };
    if (body.dense_similarity_top_k !== undefined) params.DenseSimilarityTopK = Number(body.dense_similarity_top_k);
    if (body.sparse_similarity_top_k !== undefined) params.SparseSimilarityTopK = Number(body.sparse_similarity_top_k);
    if (body.enable_reranking !== undefined) params.EnableReranking = Boolean(body.enable_reranking);
    if (body.enable_rewrite !== undefined) params.EnableRewrite = Boolean(body.enable_rewrite);
    if (body.rerank_min_score !== undefined) params.RerankMinScore = Number(body.rerank_min_score);
    if (body.rerank_top_n !== undefined) params.RerankTopN = Number(body.rerank_top_n);

    const result = await client.retrieve(params as any);
    res.json(result);
  } catch (error) {
    res.status(500).json({ error: (error as Error).message });
  }
});

app.listen(port, '0.0.0.0', () => {
  console.log(`Bailian Knowledge Base HTTP API listening on port ${port}`);
});
