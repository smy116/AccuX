using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using AccuX.Core.Cells;
using AccuX.Core.Commands;
using AccuX.Core.Configuration;
using AccuX.Core.Logging;
using AccuX.Core.Modules;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.Comment;
using AccuX.Modules.BasicFinance.Common;
using Xunit;

namespace AccuX.Modules.BasicFinance.Tests
{
    public class CommentCodecTests
    {
        [Fact]
        public void EncryptedComment_RoundTripsAndIsMarked()
        {
            var codec = new CommentCodec(new TestProtector());

            var encoded = codec.Encode("中文🙂\n第二行", CommentKind.Encrypted);
            var decoded = codec.Decode(new CellComment(true, encoded));

            Assert.NotEqual("中文🙂\n第二行", encoded);
            Assert.True(IsBase64(encoded));
            Assert.True(decoded.Exists);
            Assert.True(decoded.IsEncrypted);
            Assert.False(decoded.IsUnreadable);
            Assert.Equal("中文🙂\n第二行", decoded.Text);
        }

        [Fact]
        public void OrdinaryBase64Text_IsKeptAsPlainText()
        {
            var plain = Convert.ToBase64String(Encoding.UTF8.GetBytes("普通批注"));
            var decoded = new CommentCodec(new TestProtector()).Decode(new CellComment(true, plain));

            Assert.True(decoded.Exists);
            Assert.False(decoded.IsEncrypted);
            Assert.Equal(plain, decoded.Text);
        }

        [Fact]
        public void Base64TextThatOnlyContainsMagic_IsKeptAsPlainText()
        {
            var plain = Convert.ToBase64String(Encoding.ASCII.GetBytes("AccuX.Comment"));
            var decoded = new CommentCodec(new TestProtector()).Decode(new CellComment(true, plain));

            Assert.False(decoded.IsEncrypted);
            Assert.False(decoded.IsUnreadable);
            Assert.Equal(plain, decoded.Text);
        }

        [Fact]
        public void CorruptedEncryptedPayload_IsUnreadable()
        {
            var codec = new CommentCodec(new TestProtector());
            var encoded = codec.Encode("secret", CommentKind.Encrypted);
            var bytes = Convert.FromBase64String(encoded);
            bytes["AccuX.Comment".Length + 1] ^= 0x7F;

            var decoded = codec.Decode(new CellComment(true, Convert.ToBase64String(bytes)));

            Assert.True(decoded.Exists);
            Assert.True(decoded.IsEncrypted);
            Assert.True(decoded.IsUnreadable);
            Assert.Contains("无法解密", decoded.ErrorMessage);
        }

        [Fact]
        public void TruncatedMarkedPayload_IsUnreadable()
        {
            var codec = new CommentCodec(new TestProtector());
            var encoded = codec.Encode("secret", CommentKind.Encrypted);
            var bytes = Convert.FromBase64String(encoded);
            var truncated = new byte["AccuX.Comment".Length + 1];
            Buffer.BlockCopy(bytes, 0, truncated, 0, truncated.Length);

            var decoded = codec.Decode(new CellComment(true, Convert.ToBase64String(truncated)));

            Assert.True(decoded.IsUnreadable);
            Assert.True(decoded.IsEncrypted);
        }

        private static bool IsBase64(string text)
        {
            try
            {
                Convert.FromBase64String(text);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }
    }

    public class CommentTextValidatorTests
    {
        [Fact]
        public void CountsCrLfAsOneAndEmojiAsOneTextElement()
        {
            Assert.Equal(3, CommentTextValidator.CountTextElements("A\r\n🙂"));
        }

        [Fact]
        public void EnforcesExactlyOneThousandTextElements()
        {
            var within = new string('中', CommentTextValidator.MaxLength);
            var beyond = within + "a";

            Assert.True(CommentTextValidator.IsWithinLimit(within));
            Assert.False(CommentTextValidator.IsWithinLimit(beyond));
            Assert.Equal("当前1000/1000字", CommentTextValidator.BuildCounter(within));
        }
    }

    public class CommentCommandTests
    {
        [Fact]
        public void SavePlainComment_WritesEncodedTextToFixedTarget()
        {
            var host = new RecordingCommentHost { Existing = CellComment.None };
            var prompt = new RecordingCommentPrompt
            {
                Result = CommentDialogResult.Save("新批注", CommentKind.Plain)
            };
            var command = new CommentCommand(prompt, new CommentCodec(new TestProtector()));
            var result = Execute(command, host);

            Assert.True(result.Success);
            Assert.Equal("新批注", host.SavedText);
            Assert.Equal(1, host.CaptureCalls);
            Assert.Equal(1, host.ReadCalls);
            Assert.Equal(1, host.SaveCalls);
        }

        [Fact]
        public void SaveExistingComment_UsesOverwriteWrite()
        {
            var host = new RecordingCommentHost { Existing = new CellComment(true, "旧批注") };
            var prompt = new RecordingCommentPrompt
            {
                Result = CommentDialogResult.Save("新批注", CommentKind.Plain)
            };

            var result = Execute(new CommentCommand(prompt), host);

            Assert.True(result.Success);
            Assert.Equal("新批注", host.SavedText);
            Assert.Equal(1, host.SaveCalls);
        }

        [Fact]
        public void SaveEncryptedComment_StoresMarkedBase64Payload()
        {
            var host = new RecordingCommentHost { Existing = CellComment.None };
            var prompt = new RecordingCommentPrompt
            {
                Result = CommentDialogResult.Save("秘密", CommentKind.Encrypted)
            };
            var codec = new CommentCodec(new TestProtector());
            var result = Execute(new CommentCommand(prompt, codec), host);

            Assert.True(result.Success);
            var decoded = codec.Decode(new CellComment(true, host.SavedText));
            Assert.True(decoded.IsEncrypted);
            Assert.Equal("秘密", decoded.Text);
        }

        [Fact]
        public void ExistingEncryptedComment_IsDecodedForEditing()
        {
            var codec = new CommentCodec(new TestProtector());
            var host = new RecordingCommentHost
            {
                Existing = new CellComment(true, codec.Encode("原内容", CommentKind.Encrypted))
            };
            var prompt = new RecordingCommentPrompt
            {
                Result = CommentDialogResult.Cancelled()
            };

            var result = Execute(new CommentCommand(prompt, codec), host);

            Assert.True(result.Success);
            Assert.NotNull(prompt.LastExisting);
            Assert.True(prompt.LastExisting.IsEncrypted);
            Assert.Equal("原内容", prompt.LastExisting.Text);
        }

        [Fact]
        public void Cancel_DoesNotWrite()
        {
            var host = new RecordingCommentHost { Existing = new CellComment(true, "已有") };
            var prompt = new RecordingCommentPrompt { Result = CommentDialogResult.Cancelled() };

            var result = Execute(new CommentCommand(prompt, new CommentCodec(new TestProtector())), host);

            Assert.True(result.Success);
            Assert.False(result.ShowMessage);
            Assert.Equal(0, host.SaveCalls);
            Assert.Equal(0, host.DeleteCalls);
        }

        [Fact]
        public void Delete_RemovesExistingComment()
        {
            var host = new RecordingCommentHost { Existing = new CellComment(true, "已有") };
            var prompt = new RecordingCommentPrompt { Result = CommentDialogResult.Delete() };

            var result = Execute(new CommentCommand(prompt), host);

            Assert.True(result.Success);
            Assert.Equal(1, host.DeleteCalls);
            Assert.Equal(0, host.SaveCalls);
        }

        [Fact]
        public void UnreadableEncryptedComment_CannotBeOverwrittenButCanBeDeleted()
        {
            var host = new RecordingCommentHost { Existing = new CellComment(true, MarkedButInvalid()) };
            var codec = new CommentCodec(new TestProtector());
            var savePrompt = new RecordingCommentPrompt
            {
                Result = CommentDialogResult.Save("覆盖", CommentKind.Plain)
            };

            var saveResult = Execute(new CommentCommand(savePrompt, codec), host);
            Assert.False(saveResult.Success);
            Assert.Equal(0, host.SaveCalls);

            var deletePrompt = new RecordingCommentPrompt { Result = CommentDialogResult.Delete() };
            var deleteResult = Execute(new CommentCommand(deletePrompt, codec), host);
            Assert.True(deleteResult.Success);
            Assert.Equal(1, host.DeleteCalls);
        }

        [Fact]
        public void MoreThanOneThousandCharacters_IsRejectedWithoutWrite()
        {
            var host = new RecordingCommentHost { Existing = CellComment.None };
            var prompt = new RecordingCommentPrompt
            {
                Result = CommentDialogResult.Save(new string('字', 1001), CommentKind.Plain)
            };

            var result = Execute(new CommentCommand(prompt), host);

            Assert.False(result.Success);
            Assert.Contains("1000", result.Message);
            Assert.Equal(0, host.SaveCalls);
        }

        private static CommandResult Execute(CommentCommand command, RecordingCommentHost host)
        {
            var definition = command.CreateDefinition("test.module");
            return definition.Handler(new CommandExecutionContext(definition, CreateContext(host)));
        }

        private static IAccuXContext CreateContext(RecordingCommentHost host)
        {
            var rangeHost = new RecordingRangeHost();
            return new ModuleContext(
                new JsonConfigManager(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json")),
                NullLogger.Instance,
                rangeHost.Context,
                new RangeOperationPipeline(rangeHost, NullLogger.Instance),
                "1.0.0",
                null,
                host);
        }

        private static string MarkedButInvalid()
        {
            var bytes = Encoding.ASCII.GetBytes("AccuX.Comment");
            var payload = new byte[bytes.Length + 4];
            Buffer.BlockCopy(bytes, 0, payload, 0, bytes.Length);
            payload[bytes.Length] = 1;
            payload[bytes.Length + 1] = (byte)'b';
            payload[bytes.Length + 2] = (byte)'a';
            payload[bytes.Length + 3] = (byte)'d';
            return Convert.ToBase64String(payload);
        }
    }

    public class CommentWindowTests
    {
        [Fact]
        public void Constructor_DoesNotFailWhenWpfInitializesRadioButtons()
        {
            Exception failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var window = new CommentWindow(CommentContent.Empty, IntPtr.Zero);
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            Assert.Null(failure);
        }
    }

    internal sealed class TestProtector : ISecretProtector
    {
        private const string Prefix = "test:";

        public string Protect(string plainText)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(Prefix + (plainText ?? string.Empty)));
        }

        public string Unprotect(string protectedText)
        {
            var decoded = Encoding.UTF8.GetString(Convert.FromBase64String(protectedText ?? string.Empty));
            if (!decoded.StartsWith(Prefix, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("invalid test payload");
            }

            return decoded.Substring(Prefix.Length);
        }
    }

    internal sealed class RecordingCommentPrompt : ICommentPrompt
    {
        public CommentDialogResult Result { get; set; }

        public CommentContent LastExisting { get; private set; }

        public CommentDialogResult AskComment(CommentContent existing)
        {
            LastExisting = existing;
            return Result;
        }
    }

    internal sealed class RecordingCommentHost : ICellCommentHost
    {
        public CellComment Existing { get; set; } = CellComment.None;

        public string SavedText { get; private set; }

        public int CaptureCalls { get; private set; }

        public int ReadCalls { get; private set; }

        public int SaveCalls { get; private set; }

        public int DeleteCalls { get; private set; }

        public CellCommentTarget CaptureCommentTarget()
        {
            CaptureCalls++;
            return new CellCommentTarget("book", "sheet", "Sheet1", "$B$3", 3, 2);
        }

        public CellComment ReadComment(CellCommentTarget target)
        {
            ReadCalls++;
            return Existing;
        }

        public void SaveComment(CellCommentTarget target, string text)
        {
            SaveCalls++;
            SavedText = text;
            Existing = new CellComment(true, text);
        }

        public void DeleteComment(CellCommentTarget target)
        {
            DeleteCalls++;
            Existing = CellComment.None;
        }
    }

    internal sealed class RecordingRangeHost : IRangeOperationHost
    {
        public IHostContext Context { get; } = new CommentTestHostContext();

        public RangeTarget CaptureTarget()
        {
            return new RangeTarget("book", "sheet", "Sheet1", "A1", 1, 1, 1, false, false);
        }

        public RangeReadResult Read(RangeTarget target)
        {
            return new RangeReadResult(target, new List<CellData>());
        }

        public WriteCheckResult ValidateWrite(RangeTarget target, RangeWritePlan writePlan)
        {
            return WriteCheckResult.Success();
        }

        public void Write(RangeTarget target, RangeWritePlan writePlan)
        {
        }

        public IHostStateScope BeginStateScope(HostStateOptions options)
        {
            return new CommentTestScope();
        }

        public string[,] ReadNumberFormats(RangeTarget target)
        {
            return new string[1, 1];
        }
    }

    internal sealed class CommentTestHostContext : IHostContext
    {
        public HostKind HostKind => HostKind.Excel;

        public string HostVersion => "16.0";

        public IntPtr MainWindowHandle => IntPtr.Zero;

        public object Application => null;
    }

    internal sealed class CommentTestScope : IHostStateScope
    {
        public void Dispose()
        {
        }
    }
}
