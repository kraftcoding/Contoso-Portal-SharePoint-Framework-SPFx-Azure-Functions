
using Azure;
using Azure.Data.Tables;
using Azure.Identity;
using System.Security.Cryptography.X509Certificates;

namespace Contoso.Portal.Data.DAL.Helpers
{
    public sealed class DataStorageHelper
    {
        private readonly string _tableEndpoint;
        private readonly ClientCertificateCredential _credential;

        public DataStorageHelper(string clientId, string tenantId, X509Certificate2 certificate, string storageAccountName)
        {
            _tableEndpoint = $"https://{storageAccountName}.table.core.windows.net";
            _credential = new ClientCertificateCredential(tenantId, clientId, certificate);
        }

        public TableClient GetTableClient(string tableName)
        {
            return new TableClient(new Uri(_tableEndpoint), tableName, _credential);
        }

        public async Task DeleteAllEntitiesAsync(string tableName)
        {
            var entities = await QueryEntitiesAsync<TableEntity>(tableName);

            if (entities == null || !entities.Any())
                return;

            var client = GetTableClient(tableName);

            var groupedByPartitionKey = entities.GroupBy(e => e.PartitionKey);

            foreach (var group in groupedByPartitionKey)
            {
                var batch = new List<TableTransactionAction>();

                foreach (var entity in group)
                {
                    batch.Add(new TableTransactionAction(TableTransactionActionType.Delete, entity));

                    // Si el batch alcanza el límite de 100, lo enviamos
                    if (batch.Count == 100)
                    {
                        await client.SubmitTransactionAsync(batch);
                        batch.Clear();
                    }
                }

                // Enviar cualquier resto que quede en el batch
                if (batch.Count > 0)
                {
                    await client.SubmitTransactionAsync(batch);
                }
            }
        }

        public async Task UpsertEntityAsync(string tableName, TableEntity entity)
        {
            var client = GetTableClient(tableName);
            await client.UpsertEntityAsync(entity);
        }

        public async Task UpsertEntitiesAsync(string tableName, List<TableEntity> entities)
        {
            if (entities == null || entities.Count == 0)
                return;

            var client = GetTableClient(tableName);

            var groupedByPartitionKey = entities.GroupBy(e => e.PartitionKey);

            foreach (var group in groupedByPartitionKey)
            {
                var batch = new List<TableTransactionAction>();

                foreach (var entity in group)
                {
                    batch.Add(new TableTransactionAction(TableTransactionActionType.UpsertMerge, entity));

                    // Si el batch alcanza el límite de 100, lo enviamos
                    if (batch.Count == 100)
                    {
                        await client.SubmitTransactionAsync(batch);
                        batch.Clear();
                    }
                }

                // Enviar cualquier resto que quede en el batch
                if (batch.Count > 0)
                {
                    await client.SubmitTransactionAsync(batch);
                }
            }
        }

        public async Task<TableEntity?> GetEntityAsync(string tableName, string partitionKey, string rowKey)
        {
            var client = GetTableClient(tableName);
            try
            {
                var response = await client.GetEntityAsync<TableEntity>(partitionKey, rowKey);
                return response.Value;
            }
            catch (RequestFailedException ex) when (ex.Status == 404)
            {
                return null;
            }
        }

        public async Task<IEnumerable<T>> QueryEntitiesAsync<T>(string tableName, string? filter = null)
        where T : class, ITableEntity, new()
        {
            var client = GetTableClient(tableName);
            var results = new List<T>();

            await foreach (var entity in client.QueryAsync<T>(filter))
            {
                results.Add(entity);
            }

            return results;
        }

        public async Task DeleteEntityAsync(string tableName, string partitionKey, string rowKey)
        {
            var client = GetTableClient(tableName);
            await client.DeleteEntityAsync(partitionKey, rowKey);
        }

        // Note: Not all columns testsent in the DataStorage are defined in these classes; add them as needed
        public class MeetingEntity : ITableEntity
        {
            public string PartitionKey { get; set; } = DALConstants.DataStorage.PartitionKeys.Meeting;
            public string RowKey { get; set; } = string.ECNTy;
            public ETag ETag { get; set; }
            public DateTimeOffset? Timestamp { get; set; }
            public string? EstadoMeeting { get; set; }
            public DateTime? EndDate { get; set; }
            public DateTime? StartDate { get; set; }
            public string? NombreDepartment { get; set; }
            public string? Department { get; set; }
            public string? TipoDepartment { get; set; }
            public string? MeetingType { get; set; }
            public string? Title { get; set; }
        }

        public class DepartmentsContosoEntity : ITableEntity
        {
            public string PartitionKey { get; set; } = DALConstants.DataStorage.PartitionKeys.DepartmentContoso;
            public string RowKey { get; set; } = string.ECNTy;
            public ETag ETag { get; set; }
            public DateTimeOffset? Timestamp { get; set; }
            public string? Department { get; set; }
        }
    }
}