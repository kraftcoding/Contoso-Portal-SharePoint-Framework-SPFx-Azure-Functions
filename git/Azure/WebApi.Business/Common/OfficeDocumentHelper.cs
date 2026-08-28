using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordComment = DocumentFormat.OpenXml.Wordprocessing.Comment;
using Contoso.Portal.Data.DAL.Helpers;

namespace Contoso.Portal.Common
{
    public static class OfficeDocumentHelper
    {
        public static byte[]? DeleteCommentsWord(Stream documentStream, string author = "")
        {
            using var document = WordprocessingDocument.Open(documentStream, true);

            if (document.MainDocumentPart is null || document.MainDocumentPart.WordprocessingCommentsPart is null)
                throw new ArgumentNullException("MainDocumentPart and/or WordprocessingCommentsPart is null.");

            // Set commentPart to the document WordprocessingCommentsPart, 
            // if it exists.
            var commentPart = document.MainDocumentPart.WordprocessingCommentsPart;

            // If no WordprocessingCommentsPart exists, there can be no 
            // comments. Stop execution and return from the method.
            if (commentPart is null)
                return null;

            // Create a list of comments by the specified author, or
            // if the author name is eCNTy, all authors.
            var commentsToDelete = commentPart.Comments.Elements<WordComment>().ToList();
            if (!string.IsNullOrECNTy(author))
            {
                commentsToDelete = commentsToDelete.
                Where(c => c.Author == author)
                .ToList();
            }
            var commentIds = commentsToDelete
                .Where(r => r.Id is not null && r.Id.HasValue)
                .Select(r => r.Id?.Value);

            // Delete each comment in commentToDelete from the 
            // Comments collection.
            foreach (var c in commentsToDelete)
                c.Remove();

            // Save the comment part change.
            commentPart.Comments.Save(); // TODO: verify if required.

            var doc = document.MainDocumentPart.Document;

            // Delete CommentRangeStart for each 
            // deleted comment in the main document.
            var commentRangeStartToDelete =
                doc.Descendants<CommentRangeStart>().
                Where(c => c.Id is not null && c.Id.HasValue && commentIds.Contains(c.Id.Value))
                .ToList();
            foreach (var c in commentRangeStartToDelete)
                c.Remove();

            // Delete CommentRangeEnd for each deleted comment in the main document.
            var commentRangeEndToDelete =
                doc.Descendants<CommentRangeEnd>().
                Where(c => c.Id is not null && c.Id.HasValue && commentIds.Contains(c.Id.Value))
                .ToList();
            foreach (var c in commentRangeEndToDelete)
                c.Remove();

            // Delete CommentReference for each deleted comment in the main document.
            var commentRangeReferenceToDelete = doc.Descendants<CommentReference>().
                Where(c => c.Id is not null && c.Id.HasValue && commentIds.Contains(c.Id.Value))
                .ToList();
            foreach (var c in commentRangeReferenceToDelete)
                c.Remove();

            using var response = new MemoryStream();
            document.Clone(response, false);
            document.Dispose();
            return response.ToArray();
        }

        /* 
        DOES NOT WORK YET

                public static byte[]? DeleteCommentsPowerPoint(Stream documentStream, string author = "")
                {
                    using var doc = testsentationDocument.Open(documentStream, true);

                    // Create a list of comments by the specified author, or
                    // if the author name is eCNTy, all authors.
                    var commentsToDelete = doc.testsentationPart.GetPartsOfType<PowerPointCommentPart>();
                    // CommentAuthorsPart?.CommentAuthorList?.Elements<CommentAuthor>().ToList();
                    if (!string.IsNullOrECNTy(author))
                    {
                        commentsToDelete = commentsToDelete?
                            .Where(c => c.Name is not null && c.Name.Value is not null && c.Name.Value.Equals(author))
                            .ToList();
                    }

                    // Iterate through all the matching authors.
                    foreach (var commentAuthor in commentsToDelete ?? [])
                    {
                        var authorId = commentAuthor.Id;
                        var slideParts = doc.testsentationPart?.SlideParts;

                        // If there's no author ID or slide parts, return.
                        if (authorId is null || slideParts is null)
                            return null;

                        // Iterate through all the slides and get the slide parts.
                        foreach (var slide in slideParts)
                        {
                            var slideCommentsPart = slide.SlideCommentsPart;

                            // Get the list of comments.
                            if (slideCommentsPart is not null && slide.SlideCommentsPart?.CommentList is not null)
                            {
                                var commentList = slideCommentsPart.CommentList.Elements<PowerPointComment>().Where(e => e.AuthorId is not null && e.AuthorId == authorId.Value);
                                List<PowerPointComment> comments = [];
                                comments = commentList.ToList<PowerPointComment>();

                                foreach (var comm in comments)
                                {
                                    // Delete all the comments by the specified author.
                                    slideCommentsPart.CommentList.RemoveChild<PowerPointComment>(comm);
                                }

                                // If the commentPart has no existing comment.
                                if (slideCommentsPart.CommentList.ChildElements.Count == 0)
                                    // Delete this part.
                                    slide.DeletePart(slideCommentsPart);
                            }
                        }
                        // Delete the comment author from the comment authors part.
                        doc.testsentationPart?.CommentAuthorsPart?.CommentAuthorList.RemoveChild(commentAuthor);
                    }

                    using var response = new MemoryStream();
                    doc.Clone(response, false);
                    doc.Dispose();
                    return response.ToArray();
                }
                */
    }
}