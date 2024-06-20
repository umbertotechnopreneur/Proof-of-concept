using Aspose.Words.Replacing;
using Aspose.Words;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocLocationFinder.Helpers
{
    public class ReplaceAndInsertBookmark : IReplacingCallback
    {
        int i = 1;
        DocumentBuilder builder;
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // This is a Run node that contains either the beginning or the complete match.
            Node currentNode = e.MatchNode;

            if (builder == null)
                builder = new DocumentBuilder((Document)currentNode.Document);

            // The first (and may be the only) run can contain text before the match, 
            // in this case it is necessary to split the run.
            if (e.MatchOffset > 0)
                currentNode = SplitRun((Run)currentNode, e.MatchOffset);

            ArrayList runs = new ArrayList();

            // Find all runs that contain parts of the match string.
            int remainingLength = e.Match.Value.Length;
            while (
                (remainingLength > 0) &&
                (currentNode != null) &&
                (currentNode.GetText().Length <= remainingLength))
            {
                runs.Add(currentNode);
                remainingLength = remainingLength - currentNode.GetText().Length;

                // Select the next Run node. 
                // Have to loop because there could be other nodes such as BookmarkStart etc.
                do
                {
                    currentNode = currentNode.NextSibling;
                }
                while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
            }

            // Split the last run that contains the match if there is any text left.
            if ((currentNode != null) && (remainingLength > 0))
            {
                SplitRun((Run)currentNode, remainingLength);
                runs.Add(currentNode);
            }

            Run run = (Run)runs[0];
            builder.MoveTo(run);
            builder.StartBookmark("bookmark_" + i);
            builder.EndBookmark("bookmark_" + i);
            i++;

            // Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction.Replace;
        }

        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
        private static Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;
        }
    }
}
